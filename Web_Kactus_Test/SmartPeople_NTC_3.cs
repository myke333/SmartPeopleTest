using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using OpenQA.Selenium.Chrome;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Data;
using APITest;
using Keys = OpenQA.Selenium.Keys;
using System.IO;
using System.Management;
using System.Diagnostics;


namespace Web_Kactus_Test
{
    /// <summary>
    /// Descripción resumida de SmartPeople_NTC_3
    /// </summary>
    [CodedUITest]
    public class SmartPeople_NTC_3 : FuncionesVitales
    {
        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();

        public SmartPeople_NTC_3()
        {
        }
        [TestCleanup]
        public void Limpiar()
        {
            //Playback.PlaybackSettings.LoggerOverrideState = HtmlLoggerState.Disabled;

            DirectoryInfo di = new DirectoryInfo(TestContext.TestLogsDir);
            DirectoryInfo di1 = new DirectoryInfo(TestContext.TestRunResultsDirectory);
            DirectoryInfo di2 = new DirectoryInfo(TestContext.ResultsDirectory);

            int numfiles = di.GetFiles("*.png", SearchOption.AllDirectories).Length;
            int numfiles1 = di1.GetFiles("*.png", SearchOption.AllDirectories).Length;
            int numfiles2 = di2.GetFiles("*.png", SearchOption.AllDirectories).Length;

            if (numfiles > 0)
            {
                foreach (FileInfo file in di.GetFiles("*.png", SearchOption.AllDirectories))
                {
                    file.Delete();
                }
            }

            if (numfiles1 > 0)
            {
                foreach (FileInfo file1 in di1.GetFiles("*.png", SearchOption.AllDirectories))
                {
                    file1.Delete();
                }
            }

            if (numfiles2 > 0)
            {
                foreach (FileInfo file2 in di2.GetFiles("*.png", SearchOption.AllDirectories))
                {
                    file2.Delete();
                }
            }
            string Machine = Environment.MachineName;
            string wmiQuery = string.Format("SELECT Name, ProcessID  FROM Win32_Process WHERE (Name LIKE '{0}%{1}') OR (Name LIKE '{2}%{3}')", "CHROME", ".exe", "chrome", ".exe");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(wmiQuery);
            ManagementObjectCollection retObjectCollection = searcher.Get();
            foreach (ManagementObject retObject in retObjectCollection)
            {
                try
                {
                    int ID = Convert.ToInt32(retObject["ProcessID"]);
                    Process processes = Process.GetProcessById(ID, Machine);
                    processes.Kill();
                }
                catch
                {
                    break;
                }
            }
            string wmiQuery1 = string.Format("SELECT Name, ProcessID  FROM Win32_Process WHERE (Name LIKE '{0}%{1}') OR (Name LIKE '{2}%{3}')", "CHROMEDRIVER", ".exe", "chromedriver", ".exe");
            ManagementObjectSearcher searcher1 = new ManagementObjectSearcher(wmiQuery1);
            ManagementObjectCollection retObjectCollection1 = searcher1.Get();
            foreach (ManagementObject retObject1 in retObjectCollection1)
            {
                try
                {
                    int ID = Convert.ToInt32(retObject1["ProcessID"]);
                    Process processes1 = Process.GetProcessById(ID, Machine);
                    processes1.Kill();
                }
                catch
                {
                    break;
                }
            }
            //Playback.Cleanup();
        }

        [TestMethod]
        public void SmartPeople_frmNmVicapNTC()
        {
            List<string> errorMessages = new List<string>();
            List<string> errorMessagesMetodo = new List<string>();
            bool bandera = false;
            string enviroment = (Environment.MachineName);
            string[] auxtable = enviroment.Split('-');
            string TableOrder = "";
            if (auxtable.Length > 1)
            {
                TableOrder = (enviroment.Replace("-", "_")).ToUpper();
            }
            else
            {
                TableOrder = enviroment.ToUpper();
            }
            //TableOrder = "ktes1";

            DataSet OrderExecutionCase = SqlAdapter.SelectOrderExecution("T", TableOrder);
            int NumCasAgen = OrderExecutionCase.Tables[0].Rows.Count;
            if (NumCasAgen < 1)
            {
                errorMessages.Add("No hay casos en el agendamiento");
            }
            foreach (DataRow rowsi in OrderExecutionCase.Tables[0].Rows)
            {
                string plans = rowsi["plans"].ToString();
                string suite = rowsi["suite"].ToString();
                string CaseId = rowsi["CaseId"].ToString();
                string orders = rowsi["orders"].ToString();
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_3.SmartPeople_frmNmVicapNTC")
                {
                    DataSet sta = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                    string endstatus = null;
                    foreach (DataRow rowsta in sta.Tables[0].Rows)
                    {
                        endstatus = rowsta["states"].ToString();
                    }
                    if (endstatus == "True")
                    {

                        TFSData GetCasen = new TFSData(CaseId);
                        DataSet DataCase = GetCasen.GetParams();

                        foreach (DataRow rows in DataCase.Tables[0].Rows)
                        {
                            int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 8;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp", "GuardarEsp", "IdEsp", "NomApeEsp", "CotratoEsp","TelEsp", "AreaEsp",
                                "CentroCosEsp", "TipoEsp", "ViaticoEsp", "MotivoEsp", "FechaSolEsp",
                                "ObservEsp", "PasajeEsp"
                            };


                            //Confirma si existen las filas y guarda su contenido en una lista
                            int rowCounter = 0;
                            foreach (string valM in variablesMTM)
                            {
                                if (rows[valM].ToString().Length != 0 && rows[valM].ToString() != null)
                                {
                                    CamposTotalesMTM.Add(rows[valM].ToString());
                                }
                                rowCounter++;
                            }

                            if (rowCounter == (CamposTotalesMTM.Count))
                            {

                                try
                                {
                                    string database = "";
                                    if (CamposTotalesMTM[2].ToLower().Contains("ora"))
                                    {
                                        database = "ORA";
                                    }
                                    else
                                    {
                                        database = "SQL";
                                    }
                                    if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);


                                    //Login Empleado
                                    selenium.LoginApps(app, CamposTotalesMTM[0], CamposTotalesMTM[1], CamposTotalesMTM[2], file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);


                                    Thread.Sleep(1000);

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(CamposTotalesMTM[3]);
                                    Thread.Sleep(2000);

                                    ////a[contains(text(),'Detalle')]
                                    ///

                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    //string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");
                                    string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        //[Retornar] = CamposTotalesMTM[7],
                                        //[Guardar] = CamposTotalesMTM[8]

                                    };


                                    int controlCounter = 0;
                                    foreach (var nameControl in InitialCheck)
                                    {
                                        Thread.Sleep(100);
                                        if (nameControl.Key.ToString() != nameControl.Value.ToString())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[controlCounter] + " incorrecto, el esperado es: " + nameControl.Value + " y el encontrado es: " + nameControl.Key);
                                        }
                                        controlCounter++;
                                        Thread.Sleep(100);
                                    }

                                    var campos = new Dictionary<string, string>();

                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTelefono']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRmtNive']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomCcos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipSoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipVia']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMotViat']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblNroCuen']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecSoli_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KtxtObserSoli_lblTexto']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtDesComi_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblUtiPasa']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Identificación","Nombres y Apellidos","Nro. Contrato","Nº Telefonico o Nº Celular",
                                        "Area","Centro de Costo","Tipo","Tipo de Viatico","Motivo",
                                        "Fecha de Solicitud","Observaciones de la Solicitud","Utiliza Pasaje"
                                    };
                                    //Debugger.Launch();
                                    for (int i = 0; i < xpath.Count; i++)
                                    {
                                        xpath[i] = xpath[i].Replace("xpath=", "");
                                        xpath[i] = xpath[i].Replace("=", ",");
                                        xpath[i] = xpath[i].Insert(7, "contains(");
                                        xpath[i] = xpath[i].Insert(xpath[i].Length - 1, ")");
                                        campos.Add(xpath[i], descripcionSC[i]);
                                    }


                                    List<string> CamposPagina = new List<string>();
                                    List<string> CamposMTM = CamposTotalesMTM.GetRange(indexVariables, (CamposTotalesMTM.Count - indexVariables));
                                    Console.WriteLine(CamposMTM.Count);

                                    Dictionary<string, string> elementos = new Dictionary<string, string>();
                                    string campoName = " ";
                                    int counter = 0;

                                    //Ejecucion de la Funcion Campo Emergente por cada elemento del Diccionario
                                    //Debugger.Launch();

                                    foreach (var element in campos)
                                    {
                                        campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                        if (counter == 11) for (int i = 0; i < 12; i++) Keyboard.SendKeys("{DOWN}");
                                    }

                                    for (int i = 0; i < CamposPagina.Count; i++)
                                    {
                                        if (CamposPagina[i] == null)
                                        {
                                            CamposPagina[i] = " ";
                                        }

                                        if (CamposMTM[i] != CamposPagina[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + CamposMTM[i] + " y el encontrado es: " + CamposPagina[i]);
                                        }
                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click(xpath[0]);
                                    for (int i = 0; i < 14; i++) Keyboard.SendKeys("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Keyboard.SendKeys("{ENTER}");
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }

                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(500);
                                    Keyboard.SendKeys("{ENTER}");
                                    Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']
                                    selenium.Screenshot("Campos Necesarios", true, file);


                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
                                    if (errorMessagesMetodo.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorMessagesMetodo);
                                        Assert.Fail(string.Format("El caso de prueba presento los siguientes errores:{0}{1}",
                                                            Environment.NewLine, errorMessageString));
                                    }
                                    Thread.Sleep(3000);
                                    bandera = true;

                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();
                                        int StCount = Int32.Parse(SthCount);

                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(1000);
                                    bandera = true;
                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();

                                        int StCount = Int32.Parse(SthCount);
                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                else
                {
                    errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: El nombre de la automatizacion no corresponde");
                }
            }
            if (bandera == false)
            {
                if (errorMessages.Count > 0)
                {
                    var separator = string.Format("{0}{0}", Environment.NewLine);
                    string errorMessageString = string.Join(separator, errorMessages);

                    Assert.Inconclusive(string.Format("Las condiciones de ejecucion del caso son:{0}{1}",
                                 Environment.NewLine, errorMessageString));
                }
            }
        }

        [TestMethod]
        public void SmartPeople_frmRHFdForDesaNTC()
        {
            List<string> errorMessages = new List<string>();
            List<string> errorMessagesMetodo = new List<string>();
            bool bandera = false;
            string enviroment = (Environment.MachineName);
            string[] auxtable = enviroment.Split('-');
            string TableOrder = "";
            if (auxtable.Length > 1)
            {
                TableOrder = (enviroment.Replace("-", "_")).ToUpper();
            }
            else
            {
                TableOrder = enviroment.ToUpper();
            }
            //TableOrder = "ktes1";

            DataSet OrderExecutionCase = SqlAdapter.SelectOrderExecution("T", TableOrder);
            int NumCasAgen = OrderExecutionCase.Tables[0].Rows.Count;
            if (NumCasAgen < 1)
            {
                errorMessages.Add("No hay casos en el agendamiento");
            }
            foreach (DataRow rowsi in OrderExecutionCase.Tables[0].Rows)
            {
                string plans = rowsi["plans"].ToString();
                string suite = rowsi["suite"].ToString();
                string CaseId = rowsi["CaseId"].ToString();
                string orders = rowsi["orders"].ToString();
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_3.SmartPeople_frmRHFdForDesaNTC")
                {
                    DataSet sta = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                    string endstatus = null;
                    foreach (DataRow rowsta in sta.Tables[0].Rows)
                    {
                        endstatus = rowsta["states"].ToString();
                    }
                    if (endstatus == "True")
                    {

                        TFSData GetCasen = new TFSData(CaseId);
                        DataSet DataCase = GetCasen.GetParams();

                        foreach (DataRow rows in DataCase.Tables[0].Rows)
                        {
                            int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 7;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp", "CodEmpreEsp", "PlanCodEsp", "PlanNomEsp","ProgCodEsp", "ProgNomEsp",
                                "CurCodEsp", "CurNomEsp", "CencoCodEsp", "CencoNomEsp", "PCodEsp", "PNomEsp",
                                "NivCodEsp", "NivNomEsp", "FechaIniEsp", "FechaFinEsp"//, "RepGenEsp", "RepGraEsp"
                            };


                            //Confirma si existen las filas y guarda su contenido en una lista
                            int rowCounter = 0;
                            foreach (string valM in variablesMTM)
                            {
                                if (rows[valM].ToString().Length != 0 && rows[valM].ToString() != null)
                                {
                                    CamposTotalesMTM.Add(rows[valM].ToString());
                                }
                                rowCounter++;
                            }

                            if (rowCounter == (CamposTotalesMTM.Count))
                            {

                                try
                                {
                                    string database = "";
                                    if (CamposTotalesMTM[2].ToLower().Contains("ora"))
                                    {
                                        database = "ORA";
                                    }
                                    else
                                    {
                                        database = "SQL";
                                    }
                                    if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);


                                    //Login Empleado
                                    selenium.LoginApps(app, CamposTotalesMTM[0], CamposTotalesMTM[1], CamposTotalesMTM[2], file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    //rh
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol RRHH')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'GESTION HUMANA')]");
                                    }
                                    selenium.Screenshot("RH", true, file);

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(CamposTotalesMTM[3]);
                                    Thread.Sleep(2000);

                                    ////a[contains(text(),'Detalle')]
                                    ///

                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    //string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");
                                    //string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        //[Retornar] = CamposTotalesMTM[7],
                                        //[Guardar] = CamposTotalesMTM[8]

                                    };


                                    int controlCounter = 0;
                                    foreach (var nameControl in InitialCheck)
                                    {
                                        Thread.Sleep(100);
                                        if (nameControl.Key.ToString() != nameControl.Value.ToString())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[controlCounter] + " incorrecto, el esperado es: " + nameControl.Value + " y el encontrado es: " + nameControl.Key);
                                        }
                                        controlCounter++;
                                        Thread.Sleep(100);
                                    }

                                    var campos = new Dictionary<string, string>();

                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorPlanCodigo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorPlanNombre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorProgramaCodigo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorProgramaNombre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorCursoCodigo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorCursoNombre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorCentroCostoCodigo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCentroCostoNombre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorPlanCodigo8']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorPlanNombre8']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorNivelCargoCodigo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorNivelCargoNombre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaInicial_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaFinal_lblFecha']",
                                        //"xpath=//input[@id='ctl00_ContenidoPagina_chkIndiRep001']",
                                        //"xpath=//input[@id='ctl00_ContenidoPagina_chkIndiRep002']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Empresa","Plan Código","Plan Nombre","Progama Código","Programa Nombre",
                                        "Curso Código","Curso Nombre","Centro de Costo Código","Cento de Costo Nombre",
                                        "Plan 2 Código","Plan 2 Nombre","Nivel Código","Nivel Nombre","Fecha Inicial",
                                        "Fecha Final"//,"Reporte Resumen General","Reporte Gráfico Comparativo"
                                    };

                                    for (int i = 0; i < xpath.Count; i++)
                                    {
                                        xpath[i] = xpath[i].Replace("xpath=", "");
                                        xpath[i] = xpath[i].Replace("=", ",");
                                        xpath[i] = xpath[i].Insert(7, "contains(");
                                        xpath[i] = xpath[i].Insert(xpath[i].Length - 1, ")");
                                        campos.Add(xpath[i], descripcionSC[i]);
                                    }


                                    List<string> CamposPagina = new List<string>();
                                    List<string> CamposMTM = CamposTotalesMTM.GetRange(indexVariables, (CamposTotalesMTM.Count - indexVariables));
                                    Console.WriteLine(CamposMTM.Count);

                                    Dictionary<string, string> elementos = new Dictionary<string, string>();
                                    string campoName = " ";
                                    int counter = 0;

                                    //Ejecucion de la Funcion Campo Emergente por cada elemento del Diccionario
                                    foreach (var element in campos)
                                    {
                                        campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                        if (counter == 13) for (int i = 0; i < 10; i++) Keyboard.SendKeys("{DOWN}");

                                    }

                                    for (int i = 0; i < CamposPagina.Count; i++)
                                    {
                                        if (CamposPagina[i] == null)
                                        {
                                            CamposPagina[i] = " ";
                                        }

                                        if (CamposMTM[i] != CamposPagina[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + CamposMTM[i] + " y el encontrado es: " + CamposPagina[i]);
                                        }
                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click(xpath[0]);
                                    for (int i = 0; i < 10; i++) Keyboard.SendKeys("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Keyboard.SendKeys("{ENTER}");
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }


                                    //if (elementListPagina2.Count > 0)
                                    //{
                                    //    foreach (IWebElement pageEle in elementListPagina2)
                                    //    {
                                    //        if (pageEle.TagName == "span" && pageEle.Displayed && pageEle.GetAttribute("Class") == "rfv")
                                    //        {
                                    //            String campo = pageEle.GetAttribute("id");
                                    //            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El control de ID " + campo + " es requerido");
                                    //        }
                                    //    }
                                    //}

                                    //selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    //Thread.Sleep(500);
                                    //Keyboard.SendKeys("{ENTER}");
                                    //Thread.Sleep(500);
                                    ////xpath =//a[@id='btnGuardar']
                                    //selenium.Screenshot("Campos Necesarios", true, file);






                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
                                    if (errorMessagesMetodo.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorMessagesMetodo);
                                        Assert.Fail(string.Format("El caso de prueba presento los siguientes errores:{0}{1}",
                                                            Environment.NewLine, errorMessageString));
                                    }
                                    Thread.Sleep(3000);
                                    bandera = true;

                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();
                                        int StCount = Int32.Parse(SthCount);

                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(1000);
                                    bandera = true;
                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();

                                        int StCount = Int32.Parse(SthCount);
                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                else
                {
                    errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: El nombre de la automatizacion no corresponde");
                }
            }
            if (bandera == false)
            {
                if (errorMessages.Count > 0)
                {
                    var separator = string.Format("{0}{0}", Environment.NewLine);
                    string errorMessageString = string.Join(separator, errorMessages);

                    Assert.Inconclusive(string.Format("Las condiciones de ejecucion del caso son:{0}{1}",
                                 Environment.NewLine, errorMessageString));
                }
            }
        }

        [TestMethod]
        public void SmartPeople_frmRHNmLiqViENTC()
        {
            List<string> errorMessages = new List<string>();
            List<string> errorMessagesMetodo = new List<string>();
            bool bandera = false;
            string enviroment = (Environment.MachineName);
            string[] auxtable = enviroment.Split('-');
            string TableOrder = "";
            if (auxtable.Length > 1)
            {
                TableOrder = (enviroment.Replace("-", "_")).ToUpper();
            }
            else
            {
                TableOrder = enviroment.ToUpper();
            }
            //TableOrder = "ktes1";

            DataSet OrderExecutionCase = SqlAdapter.SelectOrderExecution("T", TableOrder);
            int NumCasAgen = OrderExecutionCase.Tables[0].Rows.Count;
            if (NumCasAgen < 1)
            {
                errorMessages.Add("No hay casos en el agendamiento");
            }
            foreach (DataRow rowsi in OrderExecutionCase.Tables[0].Rows)
            {
                string plans = rowsi["plans"].ToString();
                string suite = rowsi["suite"].ToString();
                string CaseId = rowsi["CaseId"].ToString();
                string orders = rowsi["orders"].ToString();
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_3.SmartPeople_frmRHNmLiqViENTC")
                {
                    DataSet sta = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                    string endstatus = null;
                    foreach (DataRow rowsta in sta.Tables[0].Rows)
                    {
                        endstatus = rowsta["states"].ToString();
                    }
                    if (endstatus == "True")
                    {

                        TFSData GetCasen = new TFSData(CaseId);
                        DataSet DataCase = GetCasen.GetParams();

                        foreach (DataRow rows in DataCase.Tables[0].Rows)
                        {
                            int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 7;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp", "ContEsp", "IdEsp",
                                "NomApEsp", "CargoEsp", "DepartEsp", "LiquiEsp", "TipZonEsp", "FechaSolEs",
                                "FechaSalEsp", "HoraSalEsp", "FechaRegEsp", "HoraRegEsp", "MotivoEsp", "RangEsp",
                                "ConcepEsp", "HayAntiEsp", "ValAntiEsp", "TipoEsp", "DescripEsp", "ValEsp",
                                "TotalAliEsp", "TotalTranEsp","TotalHospEsp"
                            };


                            //Confirma si existen las filas y guarda su contenido en una lista
                            int rowCounter = 0;
                            foreach (string valM in variablesMTM)
                            {
                                if (rows[valM].ToString().Length != 0 && rows[valM].ToString() != null)
                                {
                                    CamposTotalesMTM.Add(rows[valM].ToString());
                                }
                                rowCounter++;
                            }

                            if (rowCounter == (CamposTotalesMTM.Count))
                            {

                                try
                                {
                                    string database = "";
                                    if (CamposTotalesMTM[2].ToLower().Contains("ora"))
                                    {
                                        database = "ORA";
                                    }
                                    else
                                    {
                                        database = "SQL";
                                    }
                                    if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);


                                    //Login Empleado
                                    selenium.LoginApps(app, CamposTotalesMTM[0], CamposTotalesMTM[1], CamposTotalesMTM[2], file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    //rh
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol RRHH')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'GESTION HUMANA')]");
                                    }
                                    selenium.Screenshot("RH", true, file);

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(CamposTotalesMTM[3]);
                                    Thread.Sleep(2000);

                                    ////a[contains(text(),'Detalle')]
                                    ///

                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    //string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");
                                    //string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        //[Retornar] = CamposTotalesMTM[7],
                                        //[Guardar] = CamposTotalesMTM[8]

                                    };


                                    int controlCounter = 0;
                                    foreach (var nameControl in InitialCheck)
                                    {
                                        Thread.Sleep(100);
                                        if (nameControl.Key.ToString() != nameControl.Value.ToString())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[controlCounter] + " incorrecto, el esperado es: " + nameControl.Value + " y el encontrado es: " + nameControl.Key);
                                        }
                                        controlCounter++;
                                        Thread.Sleep(100);
                                    }

                                    var campos = new Dictionary<string, string>();

                                    List<string> xpath = new List<string>() {

                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCar_Empl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDep_Empl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label6']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTip_Zona']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaSolicitud_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaSalida_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlHoraSali_lblHora']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaRegreso_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlHoraRegr_lblHora']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMot_Viaje']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRangos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label5']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblHay_anti']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAnticipo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina2_lblTip_Fact']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina2_lblDes_Fact']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina2_lblVal_Fact']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina2_Label2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina2_Label3']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina2_Label4']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Contrato",
                                        "Identificacion","Nombres y Apellidos","Cargo","Departamento","Liquidación Viaticos al",
                                        "Tipo Zona","Fecha de Solicitud","Fecha de Salida","Hora Salida","Fecha de Regreso",
                                        "Hora Regreso","Motivo del Viaje","Rangos de Tiempo","Concepto","Hay Anticipo","Valor Anticipo",
                                        "Tipo","Descripción","Valor","Total Facturas Alimentación","Total Facturas Transporte","Total Facturas Hospedaje"
                                    };

                                    for (int i = 0; i < xpath.Count; i++)
                                    {
                                        xpath[i] = xpath[i].Replace("xpath=", "");
                                        xpath[i] = xpath[i].Replace("=", ",");
                                        xpath[i] = xpath[i].Insert(7, "contains(");
                                        xpath[i] = xpath[i].Insert(xpath[i].Length - 1, ")");
                                        campos.Add(xpath[i], descripcionSC[i]);
                                    }


                                    List<string> CamposPagina = new List<string>();
                                    List<string> CamposMTM = CamposTotalesMTM.GetRange(indexVariables, (CamposTotalesMTM.Count - indexVariables));
                                    Console.WriteLine(CamposMTM.Count);

                                    Dictionary<string, string> elementos = new Dictionary<string, string>();
                                    string campoName = " ";
                                    int counter = 0;

                                    //Ejecucion de la Funcion Campo Emergente por cada elemento del Diccionario
                                    foreach (var element in campos)
                                    {
                                        campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                        if (counter == 10) for (int i = 0; i < 15; i++) Keyboard.SendKeys("{DOWN}");

                                    }

                                    for (int i = 0; i < CamposPagina.Count; i++)
                                    {
                                        if (CamposPagina[i] == null)
                                        {
                                            CamposPagina[i] = " ";
                                        }

                                        if (CamposMTM[i] != CamposPagina[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + CamposMTM[i] + " y el encontrado es: " + CamposPagina[i]);
                                        }
                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click(xpath[0]);
                                    for (int i = 0; i < 20; i++) Keyboard.SendKeys("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Keyboard.SendKeys("{ENTER}");
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }


                                    //if (elementListPagina2.Count > 0)
                                    //{
                                    //    foreach (IWebElement pageEle in elementListPagina2)
                                    //    {
                                    //        if (pageEle.TagName == "span" && pageEle.Displayed && pageEle.GetAttribute("Class") == "rfv")
                                    //        {
                                    //            String campo = pageEle.GetAttribute("id");
                                    //            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El control de ID " + campo + " es requerido");
                                    //        }
                                    //    }
                                    //}

                                    //selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    //Thread.Sleep(500);
                                    //Keyboard.SendKeys("{ENTER}");
                                    //Thread.Sleep(500);
                                    ////xpath =//a[@id='btnGuardar']
                                    //selenium.Screenshot("Campos Necesarios", true, file);






                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
                                    if (errorMessagesMetodo.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorMessagesMetodo);
                                        Assert.Fail(string.Format("El caso de prueba presento los siguientes errores:{0}{1}",
                                                            Environment.NewLine, errorMessageString));
                                    }
                                    Thread.Sleep(3000);
                                    bandera = true;

                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();
                                        int StCount = Int32.Parse(SthCount);

                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(1000);
                                    bandera = true;
                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();

                                        int StCount = Int32.Parse(SthCount);
                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                else
                {
                    errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: El nombre de la automatizacion no corresponde");
                }
            }
            if (bandera == false)
            {
                if (errorMessages.Count > 0)
                {
                    var separator = string.Format("{0}{0}", Environment.NewLine);
                    string errorMessageString = string.Join(separator, errorMessages);

                    Assert.Inconclusive(string.Format("Las condiciones de ejecucion del caso son:{0}{1}",
                                 Environment.NewLine, errorMessageString));
                }
            }
        }


        [TestMethod]
        public void SmartPeople_frmRHSLRepSLNTC()
        {
            List<string> errorMessages = new List<string>();
            List<string> errorMessagesMetodo = new List<string>();
            bool bandera = false;
            string enviroment = (Environment.MachineName);
            string[] auxtable = enviroment.Split('-');
            string TableOrder = "";
            if (auxtable.Length > 1)
            {
                TableOrder = (enviroment.Replace("-", "_")).ToUpper();
            }
            else
            {
                TableOrder = enviroment.ToUpper();
            }
            //TableOrder = "ktes1";

            DataSet OrderExecutionCase = SqlAdapter.SelectOrderExecution("T", TableOrder);
            int NumCasAgen = OrderExecutionCase.Tables[0].Rows.Count;
            if (NumCasAgen < 1)
            {
                errorMessages.Add("No hay casos en el agendamiento");
            }
            foreach (DataRow rowsi in OrderExecutionCase.Tables[0].Rows)
            {
                string plans = rowsi["plans"].ToString();
                string suite = rowsi["suite"].ToString();
                string CaseId = rowsi["CaseId"].ToString();
                string orders = rowsi["orders"].ToString();
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_3.SmartPeople_frmRHSLRepSLNTC")
                {
                    DataSet sta = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                    string endstatus = null;
                    foreach (DataRow rowsta in sta.Tables[0].Rows)
                    {
                        endstatus = rowsta["states"].ToString();
                    }
                    if (endstatus == "True")
                    {

                        TFSData GetCasen = new TFSData(CaseId);
                        DataSet DataCase = GetCasen.GetParams();

                        foreach (DataRow rows in DataCase.Tables[0].Rows)
                        {
                            int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 7;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp", "EmpreEsp", "HojaEsp",  "CargEsp",
                                "AreaEsp", "ExperEsp", "InstituEsp", "EducaEsp", "PalClaEsp", "ProyecEsp",
                                "NoFormEsp", "TimeEsp", "AreaInteEsp", "GeneroEsp", "AspirEsp", "IdiomaEsp",
                                "IvesEsp", "CityEsp"
                            };


                            //Confirma si existen las filas y guarda su contenido en una lista
                            int rowCounter = 0;
                            foreach (string valM in variablesMTM)
                            {
                                if (rows[valM].ToString().Length != 0 && rows[valM].ToString() != null)
                                {
                                    CamposTotalesMTM.Add(rows[valM].ToString());
                                }
                                rowCounter++;
                            }

                            if (rowCounter == (CamposTotalesMTM.Count))
                            {

                                try
                                {
                                    string database = "";
                                    if (CamposTotalesMTM[2].ToLower().Contains("ora"))
                                    {
                                        database = "ORA";
                                    }
                                    else
                                    {
                                        database = "SQL";
                                    }
                                    if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);


                                    //Login Empleado
                                    selenium.LoginApps(app, CamposTotalesMTM[0], CamposTotalesMTM[1], CamposTotalesMTM[2], file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    //rh
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol RRHH')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'GESTION HUMANA')]");
                                    }
                                    selenium.Screenshot("RH", true, file);

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(CamposTotalesMTM[3]);
                                    Thread.Sleep(2000);

                                    ////a[contains(text(),'Detalle')]
                                    ///

                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    //string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");
                                    //string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        //[Retornar] = CamposTotalesMTM[7],
                                        //[Guardar] = CamposTotalesMTM[8]

                                    };


                                    int controlCounter = 0;
                                    foreach (var nameControl in InitialCheck)
                                    {
                                        Thread.Sleep(100);
                                        if (nameControl.Key.ToString() != nameControl.Value.ToString())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[controlCounter] + " incorrecto, el esperado es: " + nameControl.Value + " y el encontrado es: " + nameControl.Key);
                                        }
                                        controlCounter++;
                                        Thread.Sleep(100);
                                    }

                                    var campos = new Dictionary<string, string>();

                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblOrigen']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblConsuCodCarg']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblConsuNomCargo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsuCar']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAreaExp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAnosExp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblInstituciones']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEducacion']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPalabraCla']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblProyecto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEducNoFormal']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblHorasCapa']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAreaInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblGenero']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAspSalarial']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblIdioma']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTrayectoria']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrDivipUbi_lblDivPoli']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Empresa","Origen hojas de vida",
                                        "Cargo","Área de experiencia","Años de experiencia","Institución","Educación",
                                        "Palabra clave","Proyecto","Educación no formal","Tiempo de estudio","Área de interés",
                                        "Género","Aspiración salarial","Idioma","Trayectoria investigativa","Ciudad"
                                    };

                                    for (int i = 0; i < xpath.Count; i++)
                                    {
                                        xpath[i] = xpath[i].Replace("xpath=", "");
                                        xpath[i] = xpath[i].Replace("=", ",");
                                        xpath[i] = xpath[i].Insert(7, "contains(");
                                        xpath[i] = xpath[i].Insert(xpath[i].Length - 1, ")");
                                        campos.Add(xpath[i], descripcionSC[i]);
                                    }


                                    List<string> CamposPagina = new List<string>();
                                    List<string> CamposMTM = CamposTotalesMTM.GetRange(indexVariables, (CamposTotalesMTM.Count - indexVariables));
                                    Console.WriteLine(CamposMTM.Count);

                                    Dictionary<string, string> elementos = new Dictionary<string, string>();
                                    string campoName = " ";
                                    int counter = 0;

                                    //Ejecucion de la Funcion Campo Emergente por cada elemento del Diccionario
                                    foreach (var element in campos)
                                    {
                                        campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                        if (counter == 4) for (int i = 0; i < 4; i++) Keyboard.SendKeys("{DOWN}");
                                    }

                                    for (int i = 0; i < CamposPagina.Count; i++)
                                    {
                                        if (CamposPagina[i] == null)
                                        {
                                            CamposPagina[i] = " ";
                                        }

                                        if (CamposMTM[i] != CamposPagina[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + CamposMTM[i] + " y el encontrado es: " + CamposPagina[i]);
                                        }
                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click(xpath[0]);
                                    for (int i = 0; i < 4; i++) Keyboard.SendKeys("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        //Keyboard.SendKeys("{ENTER}");
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }


                                    //if (elementListPagina2.Count > 0)
                                    //{
                                    //    foreach (IWebElement pageEle in elementListPagina2)
                                    //    {
                                    //        if (pageEle.TagName == "span" && pageEle.Displayed && pageEle.GetAttribute("Class") == "rfv")
                                    //        {
                                    //            String campo = pageEle.GetAttribute("id");
                                    //            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El control de ID " + campo + " es requerido");
                                    //        }
                                    //    }
                                    //}

                                    //selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    //Thread.Sleep(500);
                                    //Keyboard.SendKeys("{ENTER}");
                                    //Thread.Sleep(500);
                                    ////xpath =//a[@id='btnGuardar']
                                    //selenium.Screenshot("Campos Necesarios", true, file);






                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
                                    if (errorMessagesMetodo.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorMessagesMetodo);
                                        Assert.Fail(string.Format("El caso de prueba presento los siguientes errores:{0}{1}",
                                                            Environment.NewLine, errorMessageString));
                                    }
                                    Thread.Sleep(3000);
                                    bandera = true;

                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();
                                        int StCount = Int32.Parse(SthCount);

                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(1000);
                                    bandera = true;
                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();

                                        int StCount = Int32.Parse(SthCount);
                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                else
                {
                    errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: El nombre de la automatizacion no corresponde");
                }
            }
            if (bandera == false)
            {
                if (errorMessages.Count > 0)
                {
                    var separator = string.Format("{0}{0}", Environment.NewLine);
                    string errorMessageString = string.Join(separator, errorMessages);

                    Assert.Inconclusive(string.Format("Las condiciones de ejecucion del caso son:{0}{1}",
                                 Environment.NewLine, errorMessageString));
                }
            }
        }
        [TestMethod]
        public void SmartPeople_frmRHSoEpactNTC()
        {
            List<string> errorMessages = new List<string>();
            List<string> errorMessagesMetodo = new List<string>();
            bool bandera = false;
            string enviroment = (Environment.MachineName);
            string[] auxtable = enviroment.Split('-');
            string TableOrder = "";
            if (auxtable.Length > 1)
            {
                TableOrder = (enviroment.Replace("-", "_")).ToUpper();
            }
            else
            {
                TableOrder = enviroment.ToUpper();
            }
            //TableOrder = "ktes1";

            DataSet OrderExecutionCase = SqlAdapter.SelectOrderExecution("T", TableOrder);
            int NumCasAgen = OrderExecutionCase.Tables[0].Rows.Count;
            if (NumCasAgen < 1)
            {
                errorMessages.Add("No hay casos en el agendamiento");
            }
            foreach (DataRow rowsi in OrderExecutionCase.Tables[0].Rows)
            {
                string plans = rowsi["plans"].ToString();
                string suite = rowsi["suite"].ToString();
                string CaseId = rowsi["CaseId"].ToString();
                string orders = rowsi["orders"].ToString();
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_3.SmartPeople_frmRHSoEpactNTC")
                {
                    DataSet sta = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                    string endstatus = null;
                    foreach (DataRow rowsta in sta.Tables[0].Rows)
                    {
                        endstatus = rowsta["states"].ToString();
                    }
                    if (endstatus == "True")
                    {

                        TFSData GetCasen = new TFSData(CaseId);
                        DataSet DataCase = GetCasen.GetParams();

                        foreach (DataRow rows in DataCase.Tables[0].Rows)
                        {
                            int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 7;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp", "ConsulIdEsp", "ConsulNameEsp", "ConsulLastEsp","ContraEsp", "NameEsp",
                                "AgeEsp", "CodIntEsp", "OldEsp", "SexoEsp", "CargoEsp", "EntiEsp",
                                "SucurEsp", "NombreEntEsp", "NombreSucuEsp", "IPSEps", "NumInfoEsp", "TipoDIagEsp",
                                "FechaAtepEsp","CentroWorkEsp", "PACEsp",
                                "RIskEsp", "AgenLesEsp", "MFAEsp", "ConsulCed","ConsulNomEsp", "ConsulApeEsp",
                                "CargoEmpEsp", "ResponsableEsp", "TipoLesionEsp", "ObservEsp", "ActoEsp"
                            };


                            //Confirma si existen las filas y guarda su contenido en una lista
                            int rowCounter = 0;
                            foreach (string valM in variablesMTM)
                            {
                                if (rows[valM].ToString().Length != 0 && rows[valM].ToString() != null)
                                {
                                    CamposTotalesMTM.Add(rows[valM].ToString());
                                }
                                rowCounter++;
                            }

                            if (rowCounter == (CamposTotalesMTM.Count))
                            {

                                try
                                {
                                    string database = "";
                                    if (CamposTotalesMTM[2].ToLower().Contains("ora"))
                                    {
                                        database = "ORA";
                                    }
                                    else
                                    {
                                        database = "SQL";
                                    }
                                    if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);


                                    //Login Empleado
                                    selenium.LoginApps(app, CamposTotalesMTM[0], CamposTotalesMTM[1], CamposTotalesMTM[2], file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    //rh
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol RRHH')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'GESTION HUMANA')]");
                                    }
                                    selenium.Screenshot("RH", true, file);

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(CamposTotalesMTM[3]);
                                    Thread.Sleep(2000);

                                    ////a[contains(text(),'Detalle')]
                                    ///

                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    //string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");
                                    //string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        //[Retornar] = CamposTotalesMTM[7],
                                        //[Guardar] = CamposTotalesMTM[8]

                                    };


                                    int controlCounter = 0;
                                    foreach (var nameControl in InitialCheck)
                                    {
                                        Thread.Sleep(100);
                                        if (nameControl.Key.ToString() != nameControl.Value.ToString())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[controlCounter] + " incorrecto, el esperado es: " + nameControl.Value + " y el encontrado es: " + nameControl.Key);
                                        }
                                        controlCounter++;
                                        Thread.Sleep(100);
                                    }

                                    var campos = new Dictionary<string, string>();

                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsuCedEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsuNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsuAplEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblContr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label5']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEdad']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_INTE']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAntiguedad']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSEX_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CARG']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_ENTI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_SUCU']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_ENTI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_SUCU']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblIPS_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNUM_FORM']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label3']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label10']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CENP']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_PART']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_AREA']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_AGEN']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_TIPO']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsuCedEmplDet2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsuNomEmplDet2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsuAplEmplDet2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCargoDet']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRespDet']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_NATU']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblOBS_NATU']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lbCODACTO']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Consultar Por Cedula", "Consultar Por Nombres", "Consultar Por Apellidos", "Contrato",
                                        "Nombres", "Edad", "Cod. Interno", "Antigüedad", "Sexo", "Cargo","Entidad", "Sucursal",
                                        "Nombre Entidad", "Nombre Sucursal", "IPS Empleado", "Número de Informe","Tipo Diagnostico",
                                        "Fecha Atep", "Centro de Trabajo", "Parte Afectada del Cuerpo", "Area De Riesgo","Agente Lesión",
                                        "Mecanismo o Forma de Accidente", "Consultar Por Cedula", "Consultar Por Nombres",
                                        "Consultar Por Apellidos", "Cargo", "Responsable", "Tipo Lesión", "Observaciones", "Acto Inseguro"
                                    };

                                    for (int i = 0; i < xpath.Count; i++)
                                    {
                                        xpath[i] = xpath[i].Replace("xpath=", "");
                                        xpath[i] = xpath[i].Replace("=", ",");
                                        xpath[i] = xpath[i].Insert(7, "contains(");
                                        xpath[i] = xpath[i].Insert(xpath[i].Length - 1, ")");
                                        campos.Add(xpath[i], descripcionSC[i]);
                                    }


                                    List<string> CamposPagina = new List<string>();
                                    List<string> CamposMTM = CamposTotalesMTM.GetRange(indexVariables, (CamposTotalesMTM.Count - indexVariables));
                                    Console.WriteLine(CamposMTM.Count);

                                    Dictionary<string, string> elementos = new Dictionary<string, string>();
                                    string campoName = " ";
                                    int counter = 0;

                                    //Ejecucion de la Funcion Campo Emergente por cada elemento del Diccionario
                                    foreach (var element in campos)
                                    {
                                        campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                        if (counter == 9) for (int i = 0; i < 13; i++) Keyboard.SendKeys("{DOWN}");
                                        if (counter == 18) for (int i = 0; i < 14; i++) Keyboard.SendKeys("{DOWN}");
                                        if (counter == 28) for (int i = 0; i < 8; i++) Keyboard.SendKeys("{DOWN}");
                                    }

                                    for (int i = 0; i < CamposPagina.Count; i++)
                                    {
                                        if (CamposPagina[i] == null)
                                        {
                                            CamposPagina[i] = " ";
                                        }

                                        if (CamposMTM[i] != CamposPagina[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + CamposMTM[i] + " y el encontrado es: " + CamposPagina[i]);
                                        }
                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click(xpath[0]);
                                    for (int i = 0; i < 35; i++) Keyboard.SendKeys("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Keyboard.SendKeys("{ENTER}");
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }


                                    //if (elementListPagina2.Count > 0)
                                    //{
                                    //    foreach (IWebElement pageEle in elementListPagina2)
                                    //    {
                                    //        if (pageEle.TagName == "span" && pageEle.Displayed && pageEle.GetAttribute("Class") == "rfv")
                                    //        {
                                    //            String campo = pageEle.GetAttribute("id");
                                    //            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El control de ID " + campo + " es requerido");
                                    //        }
                                    //    }
                                    //}

                                    //selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    //Thread.Sleep(500);
                                    //Keyboard.SendKeys("{ENTER}");
                                    //Thread.Sleep(500);
                                    ////xpath =//a[@id='btnGuardar']
                                    //selenium.Screenshot("Campos Necesarios", true, file);






                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
                                    if (errorMessagesMetodo.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorMessagesMetodo);
                                        Assert.Fail(string.Format("El caso de prueba presento los siguientes errores:{0}{1}",
                                                            Environment.NewLine, errorMessageString));
                                    }
                                    Thread.Sleep(3000);
                                    bandera = true;

                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();
                                        int StCount = Int32.Parse(SthCount);

                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(1000);
                                    bandera = true;
                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();

                                        int StCount = Int32.Parse(SthCount);
                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                else
                {
                    errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: El nombre de la automatizacion no corresponde");
                }
            }
            if (bandera == false)
            {
                if (errorMessages.Count > 0)
                {
                    var separator = string.Format("{0}{0}", Environment.NewLine);
                    string errorMessageString = string.Join(separator, errorMessages);

                    Assert.Inconclusive(string.Format("Las condiciones de ejecucion del caso son:{0}{1}",
                                 Environment.NewLine, errorMessageString));
                }
            }
        }

        [TestMethod]
        public void SmartPeople_frmBiTrinvNTC()
        {
            List<string> errorMessages = new List<string>();
            List<string> errorMessagesMetodo = new List<string>();
            bool bandera = false;
            string enviroment = (Environment.MachineName);
            string[] auxtable = enviroment.Split('-');
            string TableOrder = "";
            if (auxtable.Length > 1)
            {
                TableOrder = (enviroment.Replace("-", "_")).ToUpper();
            }
            else
            {
                TableOrder = enviroment.ToUpper();
            }
            //TableOrder = "ktes1";

            DataSet OrderExecutionCase = SqlAdapter.SelectOrderExecution("T", TableOrder);
            int NumCasAgen = OrderExecutionCase.Tables[0].Rows.Count;
            if (NumCasAgen < 1)
            {
                errorMessages.Add("No hay casos en el agendamiento");
            }
            foreach (DataRow rowsi in OrderExecutionCase.Tables[0].Rows)
            {
                string plans = rowsi["plans"].ToString();
                string suite = rowsi["suite"].ToString();
                string CaseId = rowsi["CaseId"].ToString();
                string orders = rowsi["orders"].ToString();
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_3.SmartPeople_frmBiTrinvNTC")
                {
                    DataSet sta = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                    string endstatus = null;
                    foreach (DataRow rowsta in sta.Tables[0].Rows)
                    {
                        endstatus = rowsta["states"].ToString();
                    }
                    if (endstatus == "True")
                    {

                        TFSData GetCasen = new TFSData(CaseId);
                        DataSet DataCase = GetCasen.GetParams();

                        foreach (DataRow rows in DataCase.Tables[0].Rows)
                        {
                            int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 9;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp", "RetornarEsp", "GuardarEsp",
                                "ConsecutivoEsp", "IdentificaionEsp"
                            };


                            //Confirma si existen las filas y guarda su contenido en una lista
                            int rowCounter = 0;
                            foreach (string valM in variablesMTM)
                            {
                                if (rows[valM].ToString().Length != 0 && rows[valM].ToString() != null)
                                {
                                    CamposTotalesMTM.Add(rows[valM].ToString());
                                }
                                rowCounter++;
                            }

                            if (rowCounter == (CamposTotalesMTM.Count))
                            {

                                try
                                {
                                    string database = "";
                                    if (CamposTotalesMTM[2].ToLower().Contains("ora"))
                                    {
                                        database = "ORA";
                                    }
                                    else
                                    {
                                        database = "SQL";
                                    }
                                    if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);


                                    //Login Empleado
                                    selenium.LoginApps(app, CamposTotalesMTM[0], CamposTotalesMTM[1], CamposTotalesMTM[2], file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(CamposTotalesMTM[3]);
                                    Thread.Sleep(2000);

                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Retornar", "Guardar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");
                                    string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        [Retornar] = CamposTotalesMTM[7],
                                        [Guardar] = CamposTotalesMTM[8]

                                    };


                                    int controlCounter = 0;
                                    foreach (var nameControl in InitialCheck)
                                    {
                                        Thread.Sleep(100);
                                        if (nameControl.Key.ToString() != nameControl.Value.ToString())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[controlCounter] + " incorrecto, el esperado es: " + nameControl.Value + " y el encontrado es: " + nameControl.Key);
                                        }
                                        controlCounter++;
                                        Thread.Sleep(100);
                                    }

                                    var campos = new Dictionary<string, string>();

                                    List<string> xpath = new List<string>() {

                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRmtTrin']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Consecutivo", "Identificación"
                                    };

                                    for (int i = 0; i < xpath.Count; i++)
                                    {
                                        xpath[i] = xpath[i].Replace("xpath=", "");
                                        xpath[i] = xpath[i].Replace("=", ",");
                                        xpath[i] = xpath[i].Insert(7, "contains(");
                                        xpath[i] = xpath[i].Insert(xpath[i].Length - 1, ")");
                                        campos.Add(xpath[i], descripcionSC[i]);
                                    }


                                    List<string> CamposPagina = new List<string>();
                                    List<string> CamposMTM = CamposTotalesMTM.GetRange(indexVariables, (CamposTotalesMTM.Count - indexVariables));
                                    Console.WriteLine(CamposMTM.Count);

                                    Dictionary<string, string> elementos = new Dictionary<string, string>();
                                    string campoName = " ";
                                    int counter = 0;

                                    //Ejecucion de la Funcion Campo Emergente por cada elemento del Diccionario
                                    foreach (var element in campos)
                                    {
                                        campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                    }


                                    for (int i = 0; i < CamposPagina.Count; i++)
                                    {
                                        if (CamposPagina[i] == null)
                                        {
                                            CamposPagina[i] = " ";
                                        }

                                        if (CamposMTM[i] != CamposPagina[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + CamposMTM[i] + " y el encontrado es: " + CamposPagina[i]);
                                        }
                                    }


                                    /////////// Validación TABS ////////                                    
                                    selenium.ValTabs(file);
                                    Thread.Sleep(1000);

                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
                                    if (errorMessagesMetodo.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorMessagesMetodo);
                                        Assert.Fail(string.Format("El caso de prueba presento los siguientes errores:{0}{1}",
                                                            Environment.NewLine, errorMessageString));
                                    }
                                    Thread.Sleep(3000);
                                    bandera = true;

                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();
                                        int StCount = Int32.Parse(SthCount);

                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(1000);
                                    bandera = true;
                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();

                                        int StCount = Int32.Parse(SthCount);
                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                else
                {
                    errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: El nombre de la automatizacion no corresponde");
                }
            }
            if (bandera == false)
            {
                if (errorMessages.Count > 0)
                {
                    var separator = string.Format("{0}{0}", Environment.NewLine);
                    string errorMessageString = string.Join(separator, errorMessages);

                    Assert.Inconclusive(string.Format("Las condiciones de ejecucion del caso son:{0}{1}",
                                 Environment.NewLine, errorMessageString));
                }
            }
        }

        [TestMethod]
        public void SmartPeople_frmEdInforme360NTC()
        {
            List<string> errorMessages = new List<string>();
            List<string> errorMessagesMetodo = new List<string>();
            bool bandera = false;
            string enviroment = (Environment.MachineName);
            string[] auxtable = enviroment.Split('-');
            string TableOrder = "";
            if (auxtable.Length > 1)
            {
                TableOrder = (enviroment.Replace("-", "_")).ToUpper();
            }
            else
            {
                TableOrder = enviroment.ToUpper();
            }
            //TableOrder = "ktes1";

            DataSet OrderExecutionCase = SqlAdapter.SelectOrderExecution("T", TableOrder);
            int NumCasAgen = OrderExecutionCase.Tables[0].Rows.Count;
            if (NumCasAgen < 1)
            {
                errorMessages.Add("No hay casos en el agendamiento");
            }
            foreach (DataRow rowsi in OrderExecutionCase.Tables[0].Rows)
            {
                string plans = rowsi["plans"].ToString();
                string suite = rowsi["suite"].ToString();
                string CaseId = rowsi["CaseId"].ToString();
                string orders = rowsi["orders"].ToString();
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_3.SmartPeople_frmEdInforme360NTC")
                {
                    DataSet sta = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                    string endstatus = null;
                    foreach (DataRow rowsta in sta.Tables[0].Rows)
                    {
                        endstatus = rowsta["states"].ToString();
                    }
                    if (endstatus == "True")
                    {

                        TFSData GetCasen = new TFSData(CaseId);
                        DataSet DataCase = GetCasen.GetParams();

                        foreach (DataRow rows in DataCase.Tables[0].Rows)
                        {
                            int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 7;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp","ReportesEsp", "ReportesadiEsp"
                                //,"ConsulNameEsp", "ConsulLastEsp","ContraEsp", "NameEsp",
                                //"AgeEsp", "CodIntEsp", "OldEsp", "SexoEsp", "CargoEsp", "EntiEsp",
                                //"SucurEsp", "NombreEntEsp", "NombreSucuEsp", "IPSEps", "NumInfoEsp", "TipoDIagEsp",
                                //"FechaAtepEsp","CentroWorkEsp", "PACEsp",
                                //"RIskEsp", "AgenLesEsp", "MFAEsp", "ConsulCed","ConsulNomEsp", "ConsulApeEsp",
                                //"CargoEmpEsp", "ResponsableEsp", "TipoLesionEsp", "ObservEsp", "ActoEsp"
                            };


                            //Confirma si existen las filas y guarda su contenido en una lista
                            int rowCounter = 0;
                            foreach (string valM in variablesMTM)
                            {
                                if (rows[valM].ToString().Length != 0 && rows[valM].ToString() != null)
                                {
                                    CamposTotalesMTM.Add(rows[valM].ToString());
                                }
                                rowCounter++;
                            }

                            if (rowCounter == (CamposTotalesMTM.Count))
                            {

                                try
                                {
                                    string database = "";
                                    if (CamposTotalesMTM[2].ToLower().Contains("ora"))
                                    {
                                        database = "ORA";
                                    }
                                    else
                                    {
                                        database = "SQL";
                                    }
                                    if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);


                                    //Login Empleado
                                    selenium.LoginApps(app, CamposTotalesMTM[0], CamposTotalesMTM[1], CamposTotalesMTM[2], file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(CamposTotalesMTM[3]);
                                    Thread.Sleep(2000);
                                    //Debugger.Launch();
                                    ////a[contains(text(),'Detalle')]
                                    ///

                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    //ERROR
                                    try
                                    {
                                        if (selenium.ExistControl("//a[contains(@id,'ctl00_btnCerrar')]"))
                                        {
                                            selenium.Screenshot("Error", true, file);

                                            Thread.Sleep(500);
                                            selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                        }

                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    //string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");
                                    //string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        //[Retornar] = CamposTotalesMTM[7],
                                        //[Guardar] = CamposTotalesMTM[8]

                                    };


                                    int controlCounter = 0;
                                    foreach (var nameControl in InitialCheck)
                                    {
                                        Thread.Sleep(100);
                                        if (nameControl.Key.ToString() != nameControl.Value.ToString())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[controlCounter] + " incorrecto, el esperado es: " + nameControl.Value + " y el encontrado es: " + nameControl.Key);
                                        }
                                        controlCounter++;
                                        Thread.Sleep(100);
                                    }

                                    var campos = new Dictionary<string, string>();
                                    //selenium.Click("xpath=//td[@id='Td1']/a");
                                    //selenium.Click("xpath=//a[@id='ctl00_btnCerrar']");

                                    List<string> xpath = new List<string>() {

                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTitRepo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label3']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_Label5']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblEdad']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_INTE']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblAntiguedad']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblSEX_EMPL']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CARG']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_ENTI']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_SUCU']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Reportes", "Reportes adicionales"
                                        //,"Consultar Por Nombres", "Consultar Por Apellidos", "Contrato",
                                        //"Nombres", "Edad", "Cod. Interno", "Antigüedad", "Sexo", "Cargo","Entidad", "Sucursal",
                                        //"Nombre Entidad", "Nombre Sucursal", "IPS Empleado", "Número de Informe","Tipo Diagnostico"
                                    };

                                    for (int i = 0; i < xpath.Count; i++)
                                    {
                                        xpath[i] = xpath[i].Replace("xpath=", "");
                                        xpath[i] = xpath[i].Replace("=", ",");
                                        xpath[i] = xpath[i].Insert(7, "contains(");
                                        xpath[i] = xpath[i].Insert(xpath[i].Length - 1, ")");
                                        campos.Add(xpath[i], descripcionSC[i]);
                                    }


                                    List<string> CamposPagina = new List<string>();
                                    List<string> CamposMTM = CamposTotalesMTM.GetRange(indexVariables, (CamposTotalesMTM.Count - indexVariables));
                                    Console.WriteLine(CamposMTM.Count);

                                    Dictionary<string, string> elementos = new Dictionary<string, string>();
                                    string campoName = " ";
                                    int counter = 0;

                                    //Ejecucion de la Funcion Campo Emergente por cada elemento del Diccionario
                                    foreach (var element in campos)
                                    {
                                        campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                    }

                                    for (int i = 0; i < CamposPagina.Count; i++)
                                    {
                                        if (CamposPagina[i] == null)
                                        {
                                            CamposPagina[i] = " ";
                                        }

                                        if (CamposMTM[i] != CamposPagina[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + CamposMTM[i] + " y el encontrado es: " + CamposPagina[i]);
                                        }
                                    }


                                    ///////////// Validación TABS ////////                                    
                                    //List<IWebElement> elementList = new List<IWebElement>();
                                    ////List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    //Thread.Sleep(800);
                                    //elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
                                    ////elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    //if (elementList.Count > 0)
                                    //{
                                    //    elementList[0].Click();
                                    //    Keyboard.SendKeys("{ENTER}");
                                    //    Thread.Sleep(500);
                                    //    foreach (IWebElement pageEle in elementList)
                                    //    {
                                    //        Keyboard.SendKeys("{TAB}");
                                    //        selenium.Screenshot("TAB", true, file);

                                    //        Thread.Sleep(100);
                                    //    }
                                    //}


                                    //if (elementListPagina2.Count > 0)
                                    //{
                                    //    foreach (IWebElement pageEle in elementListPagina2)
                                    //    {
                                    //        if (pageEle.TagName == "span" && pageEle.Displayed && pageEle.GetAttribute("Class") == "rfv")
                                    //        {
                                    //            String campo = pageEle.GetAttribute("id");
                                    //            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El control de ID " + campo + " es requerido");
                                    //        }
                                    //    }
                                    //}

                                    //selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    //Thread.Sleep(500);
                                    //Keyboard.SendKeys("{ENTER}");
                                    //Thread.Sleep(500);
                                    ////xpath =//a[@id='btnGuardar']
                                    //selenium.Screenshot("Campos Necesarios", true, file);






                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
                                    if (errorMessagesMetodo.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorMessagesMetodo);
                                        Assert.Fail(string.Format("El caso de prueba presento los siguientes errores:{0}{1}",
                                                            Environment.NewLine, errorMessageString));
                                    }
                                    Thread.Sleep(3000);
                                    bandera = true;

                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();
                                        int StCount = Int32.Parse(SthCount);

                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(1000);
                                    bandera = true;
                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();

                                        int StCount = Int32.Parse(SthCount);
                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                else
                {
                    errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: El nombre de la automatizacion no corresponde");
                }
            }
            if (bandera == false)
            {
                if (errorMessages.Count > 0)
                {
                    var separator = string.Format("{0}{0}", Environment.NewLine);
                    string errorMessageString = string.Join(separator, errorMessages);

                    Assert.Inconclusive(string.Format("Las condiciones de ejecucion del caso son:{0}{1}",
                                 Environment.NewLine, errorMessageString));
                }
            }
        }


        [TestMethod]
        public void SmartPeople_frmLlBiHvintLNTC()
        {
            List<string> errorMessages = new List<string>();
            List<string> errorMessagesMetodo = new List<string>();
            bool bandera = false;
            string enviroment = (Environment.MachineName);
            string[] auxtable = enviroment.Split('-');
            string TableOrder = "";
            if (auxtable.Length > 1)
            {
                TableOrder = (enviroment.Replace("-", "_")).ToUpper();
            }
            else
            {
                TableOrder = enviroment.ToUpper();
            }
            //TableOrder = "ktes1";

            DataSet OrderExecutionCase = SqlAdapter.SelectOrderExecution("T", TableOrder);
            int NumCasAgen = OrderExecutionCase.Tables[0].Rows.Count;
            if (NumCasAgen < 1)
            {
                errorMessages.Add("No hay casos en el agendamiento");
            }
            foreach (DataRow rowsi in OrderExecutionCase.Tables[0].Rows)
            {
                string plans = rowsi["plans"].ToString();
                string suite = rowsi["suite"].ToString();
                string CaseId = rowsi["CaseId"].ToString();
                string orders = rowsi["orders"].ToString();
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_3.SmartPeople_frmLlBiHvintLNTC")
                {
                    DataSet sta = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                    string endstatus = null;
                    foreach (DataRow rowsta in sta.Tables[0].Rows)
                    {
                        endstatus = rowsta["states"].ToString();
                    }
                    if (endstatus == "True")
                    {

                        TFSData GetCasen = new TFSData(CaseId);
                        DataSet DataCase = GetCasen.GetParams();

                        foreach (DataRow rows in DataCase.Tables[0].Rows)
                        {
                            int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 8;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp"
                            };

                            //Confirma si existen las filas y guarda su contenido en una lista
                            int rowCounter = 0;
                            foreach (string valM in variablesMTM)
                            {
                                if (rows[valM].ToString().Length != 0 && rows[valM].ToString() != null)
                                {
                                    CamposTotalesMTM.Add(rows[valM].ToString());
                                }
                                rowCounter++;
                            }

                            if (rowCounter == (CamposTotalesMTM.Count))
                            {

                                try
                                {
                                    string database = "";
                                    if (CamposTotalesMTM[2].ToLower().Contains("ora"))
                                    {
                                        database = "ORA";
                                    }
                                    else
                                    {
                                        database = "SQL";
                                    }
                                    if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);


                                    //Login Empleado
                                    selenium.LoginApps(app, CamposTotalesMTM[0], CamposTotalesMTM[1], CamposTotalesMTM[2], file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);
                                    //lider
                                    if (database == "ORA")
                                    {
                                        selenium.Click("//*[@id='pLider']");
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                    }
                                    selenium.Screenshot("Lider", true, file);
                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(CamposTotalesMTM[3]);
                                    Thread.Sleep(2000);

                                    ////a[contains(text(),'Detalle')]
                                    ///

                                    //titulo de la pagina
                                    selenium.Screenshot(CamposTotalesMTM[4], true, file);

                                    //ERROR
                                    try
                                    {
                                        if (selenium.ExistControl("//a[contains(@id,'ctl00_btnCerrar')]"))
                                        {
                                            selenium.Screenshot("Error", true, file);

                                            Thread.Sleep(500);
                                            selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                        }

                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    //string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");
                                    //string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        //[Retornar] = CamposTotalesMTM[7],
                                        //[Guardar] = CamposTotalesMTM[8]

                                    };


                                    int controlCounter = 0;
                                    foreach (var nameControl in InitialCheck)
                                    {
                                        Thread.Sleep(100);
                                        if (nameControl.Key.ToString() != nameControl.Value.ToString())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[controlCounter] + " incorrecto, el esperado es: " + nameControl.Value + " y el encontrado es: " + nameControl.Key);
                                        }
                                        controlCounter++;
                                        Thread.Sleep(100);
                                    }

                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
                                    if (errorMessagesMetodo.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorMessagesMetodo);
                                        Assert.Fail(string.Format("El caso de prueba presento los siguientes errores:{0}{1}",
                                                            Environment.NewLine, errorMessageString));
                                    }
                                    Thread.Sleep(3000);
                                    bandera = true;

                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();
                                        int StCount = Int32.Parse(SthCount);

                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(1000);
                                    bandera = true;
                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();

                                        int StCount = Int32.Parse(SthCount);
                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                else
                {
                    errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: El nombre de la automatizacion no corresponde");
                }
            }
            if (bandera == false)
            {
                if (errorMessages.Count > 0)
                {
                    var separator = string.Format("{0}{0}", Environment.NewLine);
                    string errorMessageString = string.Join(separator, errorMessages);

                    Assert.Inconclusive(string.Format("Las condiciones de ejecucion del caso son:{0}{1}",
                                 Environment.NewLine, errorMessageString));
                }
            }
        }

        #region Atributos de prueba adicionales

        // Puede usar los siguientes atributos adicionales conforme escribe las pruebas:

        ////Use TestInitialize para ejecutar el código antes de ejecutar cada prueba 
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{        
        //    // Para generar código para esta prueba, seleccione "Generar código para prueba automatizada de IU" en el menú contextual y seleccione uno de los elementos de menú.
        //}

        ////Use TestCleanup para ejecutar el código después de ejecutar cada prueba
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{        
        //    // Para generar código para esta prueba, seleccione "Generar código para prueba automatizada de IU" en el menú contextual y seleccione uno de los elementos de menú.
        //}

        #endregion

        /// <summary>
        ///Obtiene o establece el contexto de las pruebas que proporciona
        ///información y funcionalidad para la serie de pruebas actual.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        private TestContext testContextInstance;

        public UIMap UIMap
        {
            get
            {
                if (this.map == null)
                {
                    this.map = new UIMap();
                }

                return this.map;
            }
        }

        private UIMap map;
    }
}
