using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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


namespace Web_Kactus_Test_V2
{
    /// <summary>
    /// Descripción resumida de SmartPeople_NTC_5
    /// </summary>
    [TestClass]
    public class SmartPeople_NTC_5 : FuncionesVitales
    {

        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();

        public SmartPeople_NTC_5()
        {
        }

        [TestMethod]
        public void Prueba_NTC()
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
            TableOrder = "ktes1";

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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_5.Prueba_NTC")
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
                            

                            //
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 7;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp", "ContAnEsp", "NuevConEsp", "ConfNuCoEsp", "FechaVenEsp"
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


                                    List<string> nameBotEn = new List<string>() { Titulo, Subtitulo, Home };

                                    for (int i = 0; i < nameBotEn.Count; i++)
                                    {
                                        if (CamposTotalesMTM[4 + i] != nameBotEn[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[i] + " incorrecto, el esperado es: " + CamposTotalesMTM[4 + i] + " y el encontrado es: " + nameBotEn[i]);
                                        }

                                        Thread.Sleep(100);
                                    }


                                    var campos = new Dictionary<string, string>();

                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblContAnte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblContNue']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblContNueConf']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecVenc_lblFecha']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Contraseña Anterior", "Nueva Contraseña", "Confirmar Nueva Contraseña",
                                        "Fecha de Vencimiento"
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
                                    //selenium.Click(xpath[0]);
                                    //for (int i = 0; i < ; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            SendKeys.Send("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }


                                    //selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    //Thread.Sleep(500);
                                    //selenium.Screenshot("Campos Necesarios", true, file);

                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']

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
        public void SmartPeople_frmLIBiEdforNTC()
        {
            // public void SmartPeople_frmBIEdforNTC()
            //System.Diagnostics.Debugger.Launch();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_5.SmartPeople_frmLIBiEdforNTC")
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
                            

                            //
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 8;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp","RetornarEsp", "ModacadEsp", "NomestEsp", "NominstEsp",
                                "CiuEsp", "FechiniEsp", "FechfinEsp", "TiemEsp", "TermEsp","EstactEsp", "EstintEsp",
                                "GraduEsp", "FechgradEsp", "PromEsp", "MatrprofEsp", "FechexpEsp"
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
                                    if (CamposTotalesMTM[2].ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://dwtfsk:8093/".ToLower())
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
                                            Thread.Sleep(100);


                                        }

                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }


                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Retornar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");


                                    List<string> nameBotEn = new List<string>() { Titulo, Subtitulo, Home, Retornar };

                                    for (int i = 0; i < nameBotEn.Count; i++)
                                    {
                                        if (CamposTotalesMTM[4 + i] != nameBotEn[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[i] + " incorrecto, el esperado es: " + CamposTotalesMTM[4 + i] + " y el encontrado es: " + nameBotEn[i]);
                                        }

                                        Thread.Sleep(100);
                                    }


                                    var campos = new Dictionary<string, string>();

                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomModi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEstu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomInst']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrDivipUbi_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecInic']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecTerm']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblUniTiem']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstTerm']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstActu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblGraDuad']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecGrad']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblProCarr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMatProf']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecExtp']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Modalidad Académica", "Nombre de los Estudios", "Nombre Institución",
                                        "Ciudad","Fecha Inicio","Fecha Final","Tiempo  Unidad",
                                        "Terminado","Estudia Actualmente","Estudio Interrumpido","Graduado",
                                        "Fecha de Grado","Promedio","Matricula Profesional","Fecha Expedición "
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
                                        try
                                        {
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }
                                        catch (Exception e)
                                        {
                                            Thread.Sleep(100);
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }

                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                        if (counter == 6) for (int i = 0; i <= 4; i++) SendKeys.Send("{DOWN}");
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
                                    for (int i = 0; i < 5; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            SendKeys.Send("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }


                                    //selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    //Thread.Sleep(500);
                                    //selenium.Screenshot("Campos Necesarios", true, file);

                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']

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
        public void SmartPeople_frmLIBiEdnfoNTC()
        {
            // public void SmartPeople_frmBIEdforNTC()
            //System.Diagnostics.Debugger.Launch();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_5.SmartPeople_frmLIBiEdnfoNTC")
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
                            

                            //
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 8;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp","RetornarEsp", "ModacadEsp", "NomestEsp", "NominstEsp",
                                "FechiniEsp", "FechfinEsp","CiuEsp", "TermEsp", "EstactEsp", "TiemEsp"
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
                                    if (CamposTotalesMTM[2].ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://dwtfsk:8093/".ToLower())
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
                                            Thread.Sleep(100);


                                        }

                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }


                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Retornar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");


                                    List<string> nameBotEn = new List<string>() { Titulo, Subtitulo, Home, Retornar };

                                    for (int i = 0; i < nameBotEn.Count; i++)
                                    {
                                        if (CamposTotalesMTM[4 + i] != nameBotEn[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[i] + " incorrecto, el esperado es: " + CamposTotalesMTM[4 + i] + " y el encontrado es: " + nameBotEn[i]);
                                        }

                                        Thread.Sleep(100);
                                    }


                                    var campos = new Dictionary<string, string>();

                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblModAcad']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEstu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomInst']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecInic']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecTerm']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrDivipLecUbi_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstTerm']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstActu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblUniTiem']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Modalidad Académica", "Nombre de los Estudios", "Nombre Institución",
                                        "Fecha Inicio","Fecha Final","Ciudad","Terminado","Estudia Actualmente",
                                        "Tiempo Estudio"
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
                                        try
                                        {
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }
                                        catch (Exception e)
                                        {
                                            Thread.Sleep(100);
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }

                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                        if (counter == 6) for (int i = 0; i < 2; i++) SendKeys.Send("{DOWN}");
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
                                    for (int i = 0; i < 2; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {

                                            SendKeys.Send("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }


                                    //selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    //Thread.Sleep(500);
                                    //selenium.Screenshot("Campos Necesarios", true, file);

                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']

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
        public void SmartPeople_frmLIBiEmplNTC()
        {
            // public void SmartPeople_frmBIEdforNTC()
            //System.Diagnostics.Debugger.Launch();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_5.SmartPeople_frmLIBiEmplNTC")
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
                            

                            //
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 8;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp","RetornarEsp", "IdenEsp","CodiEsp","NomapEsp","LugexpdocEsp","LugnacEsp",
                                "NacEsp","GenEsp","paisreEsp","DirEsp","BarEsp","TelEsp","RutEsp","TelmoEsp",
                                "TelfaEsp","MaiperEsp","EstciviEsp","LibmilEsp","NumEsp","DistEsp","GradedbasmedEsp",
                                "TitobEsp","TitprofEsp","MatprofEsp"
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
                                    if (CamposTotalesMTM[2].ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://dwtfsk:8093/".ToLower())
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

                                    //Rol Lider
                                    if (database == "ORA")
                                    {
                                        selenium.Click("//*[@id='pLider']");
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                    }


                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(CamposTotalesMTM[3]);
                                    Thread.Sleep(2000);

                                    ////a[contains(text(),'Detalle')]

                                    //titulo de la pagina
                                    selenium.Screenshot(CamposTotalesMTM[4], true, file);

                                    //ERROR
                                    try
                                    {
                                        if (selenium.ExistControl("//a[contains(@id,'ctl00_btnCerrar')]"))
                                        {
                                            selenium.Screenshot("Error", true, file);

                                            Thread.Sleep(1500);
                                            selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                            Thread.Sleep(10000);


                                        }

                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }


                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Retornar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");


                                    List<string> nameBotEn = new List<string>() { Titulo, Subtitulo, Home, Retornar };

                                    for (int i = 0; i < nameBotEn.Count; i++)
                                    {
                                        if (CamposTotalesMTM[4 + i] != nameBotEn[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[i] + " incorrecto, el esperado es: " + CamposTotalesMTM[4 + i] + " y el encontrado es: " + nameBotEn[i]);
                                        }

                                        Thread.Sleep(100);
                                    }


                                    var campos = new Dictionary<string, string>();

                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrDivipLecExpe_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrDivipLecNaci_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNacIona']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSexEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrDivipLecRes_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divDirResi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divBarResi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divTelResi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divRutResi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divTelMovi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divTelFaxi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divEeeMail']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divEstCivi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblClaLmil']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNumLmil']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDisLmil']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblGraEduc']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTitObte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divProTitu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divMatProf']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Identificación", "Código Interno", "Nombres  Apellidos",
                                        "Lugar Expedición Documento","Lugar de Nacimiento","Nacionalidad",
                                        "Genero","País de Residencia","Dirección","Barrio","Teléfono","Ruta",
                                        "Teléfono Movil","Teléfono Fax","E-Mail Personal","Estado Civil",
                                        "Libreta Militar","Número","Distrito","Grado Educación Básica y Media",
                                        "Titulo Obtenido","Titulo Profesional","Matricula Profesional",
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
                                        try
                                        {
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }
                                        catch (Exception e)
                                        {
                                            Thread.Sleep(1000);
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }

                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(2000);
                                        if (counter == 5) for (int i = 0; i < 14; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 20) for (int i = 0; i < 4; i++) SendKeys.Send("{DOWN}");

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
                                    for (int i = 0; i < 18; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            SendKeys.Send("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
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




        [TestMethod]
        public void SmartPeople_frmLIBiEmplLNTC()
        {
            // public void SmartPeople_frmBIEdforNTC()
            //System.Diagnostics.Debugger.Launch();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_5.SmartPeople_frmLIBiEmplLNTC")
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
                            

                            //
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 7;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp","ConidEsp","ConnomEsp","ConapEsp","ConcodcargEsp"
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
                                    if (CamposTotalesMTM[2].ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://dwtfsk:8093/".ToLower())
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

                                    //Rol Lider
                                    if (database == "ORA")
                                    {
                                        selenium.Click("//*[@id='pLider']");
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                    }

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(CamposTotalesMTM[3]);
                                    Thread.Sleep(2000);

                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);
                                    //Debugger.Launch();
                                    //ERROR
                                    /*try
                                    {
                                        if (selenium.ExistControl("//a[contains(@id,'ctl00_btnCerrar')]"))
                                        {
                                            selenium.Screenshot("Error", true, file);

                                            Thread.Sleep(500);
                                            selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                            Thread.Sleep(100);


                                        }

                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }
                                    */

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");


                                    List<string> nameBotEn = new List<string>() { Titulo, Subtitulo, Home };

                                    for (int i = 0; i < nameBotEn.Count; i++)
                                    {
                                        if (CamposTotalesMTM[4 + i] != nameBotEn[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[i] + " incorrecto, el esperado es: " + CamposTotalesMTM[4 + i] + " y el encontrado es: " + nameBotEn[i]);
                                        }

                                        Thread.Sleep(100);
                                    }


                                    var campos = new Dictionary<string, string>();

                                    /*List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsuCedEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsuNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsuAplEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsultarCargo']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Consultar por Identificación", "Consultar por Nombres",
                                        "Consultar por Apellidos","Consultar por código de Cargo"
                                    };*/

                                    List<string> xpath = new List<string>() {
                                        "xpath=//*[@id='ctl00_lblMenMarco']",
                                        "xpath=//*[@id='ctl00_ContenidoPagina_lblEspeciales']",
                                        "xpath=//*[@id='ctl00_ContenidoPagina_lblCodEmpr']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Datos Basicos de Mis Colaboradores", "Especiales",
                                        "Empresa"
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

                                    /*//Ejecucion de la Funcion Campo Emergente por cada elemento del Diccionario
                                    foreach (var element in campos)
                                    {
                                        try
                                        {
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }
                                        catch (Exception e)
                                        {
                                            Thread.Sleep(100);
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }

                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                        //if (counter == 5) for (int i = 0; i < 15; i++) SendKeys.Send("{DOWN}");

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
                                    }*/


                                    /////////// Validación TABS ////////
                                    /*//selenium.Click(xpath[0]);
                                    //for (int i = 0; i < 18; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            SendKeys.Send("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }*/

                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_ddlCodEmpr']");
                                    for (int i = 0; i < 3; i++)
                                    {
                                        SendKeys.Send("{TAB}");
                                        selenium.Screenshot("TAB", true, file);

                                        Thread.Sleep(100);
                                    }




                                    //selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    //Thread.Sleep(500);
                                    //selenium.Screenshot("Campos Necesarios", true, file);

                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']

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
        public void SmartPeople_frmLIBiEntprNTC()
        {
            // public void SmartPeople_frmBIEdforNTC()
            //System.Diagnostics.Debugger.Launch();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_5.SmartPeople_frmLIBiEntprNTC")
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
                            

                            //
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 9;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp","RetornarEsp", "GuardarEsp","PrendEsp","CantEsp","FechentrEsp"
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
                                    if (CamposTotalesMTM[2].ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://dwtfsk:8093/".ToLower())
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
                                            Thread.Sleep(100);


                                        }

                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }


                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Retornar", "Guardar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");
                                    string Guardar = selenium.EmergenteBotones("btnGuardar");


                                    List<string> nameBotEn = new List<string>() { Titulo, Subtitulo, Home, Retornar, Guardar };

                                    for (int i = 0; i < nameBotEn.Count; i++)
                                    {
                                        if (CamposTotalesMTM[4 + i] != nameBotEn[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[i] + " incorrecto, el esperado es: " + CamposTotalesMTM[4 + i] + " y el encontrado es: " + nameBotEn[i]);
                                        }

                                        Thread.Sleep(100);
                                    }


                                    var campos = new Dictionary<string, string>();

                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodPren']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCanEntr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecEntr_lblFecha']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Prenda", "Cantidad", "Fecha Entregada"
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
                                        try
                                        {
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }
                                        catch (Exception e)
                                        {
                                            Thread.Sleep(100);
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }

                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                        //if (counter == 6) for (int i = 0; i < 2; i++) SendKeys.Send("{DOWN}");
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
                                    //for (int i = 0; i < 2; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            SendKeys.Send("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }


                                    //selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    //Thread.Sleep(500);
                                    //selenium.Screenshot("Campos Necesarios", true, file);

                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']

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
        public void SmartPeople_frmLIBiHvextNTC()
        {
            // public void SmartPeople_frmBIEdforNTC()
            //System.Diagnostics.Debugger.Launch();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_5.SmartPeople_frmLIBiHvextNTC")
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
                            

                            //
                            //Indice desde donde empiezan los campos Emergentes //IMPORTANTE
                            int indexVariables = 8;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp","RetornarEsp", "EmpEsp", "DirEsp", "TelEsp","TipempEsp",
                                "CorrentEsp", "SaldevEsp", "EmpactEsp", "CargejeEsp","DediEsp", "FechiEsp",
                                "FechrEsp", "ManperEsp", "CargdesEsp", "AreaEsp", "TipcontrEsp",
                                "MotretEsp", "CiuEsp", "JefeiEsp", "CargjefeiEsp", "TiemservEsp",
                                "ActempEsp", "FunciEsp", "AreaexEsp"
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
                                    if (CamposTotalesMTM[2].ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (CamposTotalesMTM[2].ToLower() == "http://dwtfsk:8093/".ToLower())
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

                                    //titulo de la pagina
                                    selenium.Screenshot(CamposTotalesMTM[4], true, file);

                                    //ERROR
                                    try
                                    {
                                        if (selenium.ExistControl("//a[contains(@id,'ctl00_btnCerrar')]"))
                                        {
                                            selenium.Screenshot("Error", true, file);

                                            Thread.Sleep(1500);
                                            selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                            Thread.Sleep(10000);


                                        }

                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }


                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Retornar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");


                                    List<string> nameBotEn = new List<string>() { Titulo, Subtitulo, Home, Retornar };

                                    for (int i = 0; i < nameBotEn.Count; i++)
                                    {
                                        if (CamposTotalesMTM[4 + i] != nameBotEn[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[i] + " incorrecto, el esperado es: " + CamposTotalesMTM[4 + i] + " y el encontrado es: " + nameBotEn[i]);
                                        }

                                        Thread.Sleep(100);
                                    }


                                    var campos = new Dictionary<string, string>();

                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDirEmpr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTelEmpr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipEmpr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEntMail']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSalDemp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEmpActu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCarEjec']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDedIcac']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecIngr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecReti']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblManPers']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCarDese']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDepEmpr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMotReti']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrDivipEmpre_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblJefInme']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCarJefe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTieServ']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblActEmp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtValMFunReal_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAreExp']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Empresa", "Dirección", "Teléfono", "Tipo de Empresa",
                                        "Correo de la Entidad","Salario Devengado","Empleo Actual",
                                        "Cargo Ejecutivo","Dedicación","Fecha de Ingreso","Fecha de Retiro",
                                        "Maneja Personal","Cargo Desempeñado","Area","Tipo de Contrato",
                                        "Motivo de Retiro","Ciudad","Jefe Inmediato","Cargo del Jefe Inmediato",
                                        "Tiempo de Servicio","Actividad Empresa","Funciones","Areas de Experiencia"
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
                                        try
                                        {
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }
                                        catch (Exception e)
                                        {
                                            Thread.Sleep(1000);
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }

                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(1000);
                                        if (counter == 17) for (int i = 0; i < 13; i++) SendKeys.Send("{DOWN}");
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
                                    for (int i = 0; i < 13; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            SendKeys.Send("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }


                                    //selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    //Thread.Sleep(500);
                                    //selenium.Screenshot("Campos Necesarios", true, file);

                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']

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
        public void CodedUITestMethod1()
        {
            // Para generar código para esta prueba, seleccione "Generar código para prueba automatizada de IU" en el menú contextual y seleccione uno de los elementos de menú.
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

        
    }
}
