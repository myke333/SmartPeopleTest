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
    /// Descripción resumida de SmartPeople_NTC_4
    /// </summary>
    [TestClass]
    public class SmartPeople_NTC_4 : FuncionesVitales
    {

        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();

        public SmartPeople_NTC_4()
        {
        }

        [TestMethod]
        public void SmartPeople_frmBpBevenNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmBpBevenNTC")
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
                                "HomeEsp", "RetornarEsp", "TipProgEsp",  "NumParEsp", "NumInscEsp",
                                "ProgramaEsp", "CentCostEsp", "ListSecuEsp", "PaisEsp", "CupCompEsp", "DepartEsp", "MuniEsp", "LocalidadEsp",
                                "LugarEspecificoEsp", "DesDetallaEsp", "MisObsEsp", "TipoAsisEsp", "FormPagoEsp", "NumCuotaEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipEven']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblHorInic']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblHorFina']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNumPart']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblParInsc']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomProg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCcos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblLisSecu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPaiEven']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCupComp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDepEven']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMunEven']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblLocEven']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblLugEven']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlValMDesActi_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtOBS_ERVA_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipAsis']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblForPago']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblListSec']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Tipo de Programa",  "Nro. Participantes",
                                        "Nro. Inscritos", "Programa", "Centro de Costo", "Listado Secuencial",
                                        "País", "Cupos Compartidos", "Departamento", "Municipio", "Localidad",
                                        "Lugar Especifico", "Descripción Detallada", "Mis Observaciones",
                                        "Tipo Asistente", "Forma de Pago", "Número de Cuotas"
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
                                        if (counter == 6) for (int i = 0; i < 10; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 13) for (int i = 0; i < 10; i++) SendKeys.Send("{DOWN}");

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
                                    for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
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
        public void SmartPeople_frmBpBevenDNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmBpBevenDNTC")
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
                                "HomeEsp", "RetornarEsp", "ActividadEsp", "TipProgramEsp",
                                "ProgramEsp",
                                "PaisEsp", "LugarEspeEsp", "DepartEsp", "NumEquipoEsp", "CupCompEsp", "MuniEsp",
                                "LocalidadEsp", "TipDisEsp", "MinEsp", "MaxEsp"
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

                                            Thread.Sleep(5000);

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEvento']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipEven']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblHorInic']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblHorFina']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomProg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPaiEven']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblLugEven']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDepEven']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNumPart']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCupComp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMunEven']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblLocEven']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lbltipDisi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMinPerc']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMaxPerc']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Actividad", "Tipo de Programa",
                                        "Programa", "País", "Lugar Especifico", "Departamento", "Nro. Equipos",
                                        "Cupos Compartidos", "Municipio", "Localidad", "Tipo Disciplina", "Mínimo",
                                        "Máximo"
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
                                        if (counter == 8) for (int i = 0; i < 7; i++) SendKeys.Send("{DOWN}");
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
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 25; i++) SendKeys.Send("{UP}");
                                    //ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

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
        public void SmartPeople_frmBpSopreENTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmBpSopreENTC")
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
                                "HomeEsp", "DesPreEsp", "AdDocEsp", "NumConEsp", "IdentEsp", "NomApellEsp", "CodIntEsp",
                                "CargoEsp", "FechaIngreEsp", "TipoSalEsp", "SueldoBasEsp", "IngresosEsp", "EstadoCiEsp",
                                "FechaSolCoEsp", "NumRadEsp", "NumEsp", "PlazoAñosEsp", "ValSoliEsp", "EstadoSol", "CampReq"
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
                                            Thread.Sleep(1500);

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodPres']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAdjDocu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNmbEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCFFecIngr_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipSala']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSueBasi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTotInmt']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstCivi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecSoli_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroRadi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRmtSoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPlaPres']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValPres']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstSoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMensaje']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Descripcion Préstamo", "Adjuntat Documento Capacidad de Pago",
                                        "Nro Contrato", "Identificación", "Nombres y Apellidos", "Cod. Interno",
                                        "Cargo", "Fecha Ingreso", "Tipo de Salario", "Sueldo Básico", "Ingresos", "Estado Civil",
                                        "Fecha Solicitud y Corte", "Nro Radicación", "Número", "Plazo en Años",
                                        "Valor Solicitado", "Estado de la Solicitud", "Campos Requeridos"
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
                                        if (counter == 15) for (int i = 0; i < 7; i++) SendKeys.Send("{DOWN}");

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
                                    for (int i = 0; i < 12; i++) SendKeys.Send("{UP}");
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
        public void SmartPeople_frmBiHerraNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmBiHerraNTC")
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
                                "HomeEsp", "RetornarEsp", "GuardarEsp", "HerrEsp", "NivelEsp", "ConoEsp",
                                "ManejoEsp", "ExpeEsp", "ObsEsp",
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomHerr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNivHerr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNivCono']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNivMane']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNivExpe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_txtObsErva_lblTexto']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Herramientas", "Nivel", "Conocimiento", "Manejo",
                                        "Experiencia", "Observaciones"
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
                                    //for (int i = 0; i < 4; i++) SendKeys.Send("{UP}");
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


                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Campos Necesarios", true, file);

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
        public void SmartPeople_frmNmOpreDNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmNmOpreDNTC")
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
                                "HomeEsp", "RetornarEsp", "IdentEsp", "NomApelEsp", "NumContEsp", "CentCostEsp",
                                "MotEsp", "FechaResEsp", "HoraIniEsp", "HoraFinEsp", "JustEsp"
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
                                            Thread.Sleep(5000);

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCcos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodMhoe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecReg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblHorInic']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblHorFina']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtJustif_lblTexto']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Identificación", "Nombres y Apellidos", "Nro. de Contrato",
                                        "Centro de Costos", "Motivo", "Fecha Registro", "Hora Inicio",
                                        "Hora Final", "Justificación"
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
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 25; i++) SendKeys.Send("{UP}");
                                    //ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

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
        public void SmartPeople_frmLNmOprecNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmLNmOprecNTC")
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
                                "HomeEsp", "RetornarEsp", "GuardarEsp", "LugAdEsp", "AreaEsp", "EstadoEsp", "FechaSolEsp",
                                "FechaIniEsp", "FechaFinEsp", "IdentEsp", "NomEsp", "Ident2Esp", "Nom2Esp", "ObjGenEsp"
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
                                            Thread.Sleep(5000);

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsuCedEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConsuCedEmpl0']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstApro']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecSol_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechIni_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechFin_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lbljefInm']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomFeje']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAutor']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomAutor']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtObsGral_lblTexto']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Lugar Administrativo", "Área", "Estado", "Fecha Solicitud",
                                        "Fecha Inicio", "Fecha Fin", "Identificación", "Nombre", "Identificación",
                                        "Nombres", "Objetivos Generales"
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
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 25; i++) SendKeys.Send("{UP}");
                                    //ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

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
        public void SmartPeople_frmNmCaltuANTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmNmCaltuANTC")
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
                                "HomeEsp", "Empresa", "ProFechaInEsp", "FechaFinEsp", "FechaPagEsp", "IdentEsp",
                                "NomEsp", "ApellEsp"
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

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(CamposTotalesMTM[3]);
                                    Thread.Sleep(2000);

                                    ////a[contains(text(),'Detalle')]

                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    ////ERROR
                                    //try
                                    //{
                                    //    if (selenium.ExistControl("//a[contains(@id,'ctl00_btnCerrar')]"))
                                    //    {
                                    //        selenium.Screenshot("Error", true, file);

                                    //        Thread.Sleep(500);
                                    //        selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");


                                    //    }

                                    //}
                                    //catch (Exception e)
                                    //{
                                    //    continue;
                                    //}


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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpr']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblNom_Cargo']",

                                        "xpath=//span[@id='ctl00_ContenidoPagina_KcfFecProgIni_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KcfFecProgFin_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KcfFecPago_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFiltroId']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFiltroNombre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFiltroApellido']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Empresa","Programación Fecha Final", "Fecha Final", "Fecha Pago", "Identificación",
                                        "Nombre", "Apellido"
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
                                    //for (int i = 0; i < 4; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[2].Click();
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
        public void SmartPeople_frmBiFamNmBenLNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmBiFamNmBenLNTC")
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
                                "HomeEsp", "ContratoEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Contrato"
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
                                    selenium.Click("//span[contains(@id,'ctl00_lblTitulo')]");
                                    Thread.Sleep(500);
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$ddlNroCont')]")));
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
        public void SmartPeople_frmFdPlcurNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmFdPlcurNTC")
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
                                "HomeEsp", "ParamEsp", "SecuEsp", "CursoEsp", "FechaIniEsp", "FechaFinEsp", "CupEsp",
                                "CupAsDepEsp", "LugCurEsp", "ObjCurEsp", "ContCurEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblParametro']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSecuencial']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCurso']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecF']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCupos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCupo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblLugCurso']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrltxtObjEtiv_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrltxtConTeni_lblTexto']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Parametro", "Secuencial", "Curso", "Fecha Inicial", "Fecha Final",
                                        "Cupos", "Cupos Asignados por Dependencia", "Lugar del Curso", "Objetivo del Curso",
                                        "Contenido del Curso"
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
        public void SmartPeople_frmFdSocueNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmFdSocueNTC")
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
                                "HomeEsp", "EliminarEsp", "ActualizarEsp", "NumContEsp", "IdentEsp", "NomEsp", "CursoEsp",
                                "ProgEsp", "PlanEsp", "TipEduEsp", "ValCuEsp", "ValMonEsp", "MonedaEsp",
                                "FechaIniEsp", "PrioEsp", "CentConstEsp", "EsSolEsp"
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
                                            Thread.Sleep(5000);

                                        }

                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }


                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Eliminar", "Actualizar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Eliminar = selenium.EmergenteBotones("ctl00_btnEliminar");
                                    string Actualizar = selenium.EmergenteBotones("btnActualizar");


                                    List<string> nameBotEn = new List<string>() { Titulo, Subtitulo, Home, Eliminar, Actualizar };

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCurs']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodProg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodPlan']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipEduc']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValCurs']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValOtrm']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodMone']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecSoli_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipPrio']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCcos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstSoli']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Nro Contrato", "Identificación", "Nombres", "Curso", "Programa", "Plan",
                                        "Tipo de Educación", "Valor de Curso", "Valor Otras Monedas", "Moneda",
                                        "Fecha Inicial", "Prioridad", "Centro de Costo", "Estado Solicitud"
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
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 25; i++) SendKeys.Send("{UP}");
                                    //ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

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
        public void SmartPeople_frmGnSclavNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmGnSclavNTC")
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
                                "HomeEsp", "ConAntEsp", "NuevaConEsp", "ConfNuConEsp", "FechaVenEsp"
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
                                            Thread.Sleep(5000);

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
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 25; i++) SendKeys.Send("{UP}");
                                    //ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

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
        public void SmartPeople_frmEd3dlicNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmEd3dlicNTC")
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
                                "HomeEsp", "NumContEsp", "IdentEsp", "CodInterEsp", "NomEsp", "CargoEsp",
                                "FechaDeEsp", "FechaHaEsp", "ProtEvaEsp", "RolEsp", "DesEvaEsp", "ExEvaEsp",
                                "FortEsp", "RecoMeEsp"
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

                                            Thread.Sleep(3000);
                                            selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                            Thread.Sleep(3000);
                                            selenium.Click("//button/div/div/img");
                                            Thread.Sleep(3000);

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecDesd']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecHast']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodTeva']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodRole']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divDesEvalT']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtValMOBS_EVAL_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtValMTEX_FORT_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtValMTEX_RECO_lblTexto']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Nro. Contrato", "Identificación", "Cod. Interno", "Nombres y Apellidos",
                                        "Cargo", "Fecha Desde", "Fecha Hasta", "Prototipo de Evaluación", "Rol",
                                        "Descripción de Evaluación", "Explicación de Evaluación", "Fortalezas", "Recomendaciones de Mejoramiento"
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
                                        if (counter == 9)
                                        {
                                            selenium.Click("//h2/button/span");
                                        }
                                        if (counter == 11) for (int i = 0; i < 10; i++) SendKeys.Send("{DOWN}");
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
                                    for (int i = 0; i < 10; i++) SendKeys.Send("{UP}");
                                    Thread.Sleep(1000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_fotEmpl']");
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


                                    //selenium.Click("//a[contains(@id,'btnActualizar')]");
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
        public void SmartPeople_frmEd3dlicPPNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmEd3dlicPPNTC")
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
                                "HomeEsp", "NumContEsp", "IdentEsp", "NomEsp", "ApellEsp", "CodInterEsp",  "CargoEsp",
                                "FechaDeEsp", "FechaHaEsp", "ProtEvaEsp", "RolEsp", "DesEvaEsp", "ExEvaEsp",
                                "FortEsp", "RecoMeEsp"
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
                                            Thread.Sleep(2000);


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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label1']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNombres']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label3']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label4']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label6']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label5']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label8']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label9']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divDesEvalT']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtValMOBS_EVAL_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtValMTEX_FORT_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtValMTEX_RECO_lblTexto']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Nro. Contrato", "Identificación", "Nombres", "Apellidos", "Cod. Interno",
                                        "Cargo", "Fecha Desde", "Fecha Hasta", "Prototipo de Evaluación", "Rol",
                                        "Descripción de Evaluación", "Explicación de Evaluación", "Fortalezas", "Recomendaciones de Mejoramiento"
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

                                        if (counter == 11) for (int i = 0; i < 20; i++) SendKeys.Send("{DOWN}");

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
                                    if (database == "ORA")
                                    {
                                        for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
                                    }

                                    for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[1].Click();
                                        Thread.Sleep(500);

                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            SendKeys.Send("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }


                                    //selenium.Click("//a[contains(@id,'btnActualizar')]");
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
        public void SmartPeople_frmFdDplinNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmFdDplinNTC")
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
                                "HomeEsp", "RetornarEsp", "ParamEsp", "SecuEsp", "CursoEsp", "FechaIni", "FechaFin",
                                "CuposEsp", "NumParEsp", "SesionEsp", "FechaIn2Esp", "FechaFin2Esp", "HoraIniEsp",
                                "HoraFinEsp", "TipAsEsp", "LugSesEsp", "ContCurEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblParametro']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSecuencial']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCurso']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecInic_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecFina_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCupos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCupo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomSesi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecInicS_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecFinaS_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblHorInic']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblHorFina']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipAsis']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblLugSesi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConTeni']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Parametro", "Secuencial", "Curso", "Fecha Inicial", "Fecha Final",
                                        "Cupos", "No. Participantes Inscritos", "Sesión", "Fecha Inicial", "Fecha Final",
                                        "Hora Inicial", "Hora Final", "Tipo de Asistentes", "Lugar de la Sesión", "Contenido del Curso"
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
                                        if (counter == 12) for (int i = 0; i < 5; i++) SendKeys.Send("{DOWN}");
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


                                    //selenium.Click("//a[contains(@id,'btnActualizar')]");
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
        public void SmartPeople_frmNmAntcaDNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmNmAntcaDNTC")
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
                                "HomeEsp", "TipOpEsp", "IdentEsp", "NomApellEsp", "NumContEsp", "FechaViEsp", "HoraViEsp",
                                "FechaRegreEsp", "HoraRegreEsp", "ClienteEsp", "DetClienteEsp", "SucuEsp", "ProEsp", "AreaEsp",
                                "CentCostEsp", "ValSolEsp", "DescEsp", "ObsEsp", "ObsAntEsp", "ObsSolEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label11']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label1']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label3']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecVia_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlHorVi_lblHora']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecReg_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlHorRe_lblHora']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label15']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label16']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label6']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label5']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label4']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label9']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label10']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label8']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_OBS_ERVA_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KtxtObserSoli_lblTexto']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Tipo de Operación", "Identificación", "Nombres y Apellidos",
                                        "Nro Contrato", "Fecha de Viaje", "Hora de Viaje", "Fecha de Regreso",
                                        "Hora de Regreso", "Cliente", "Detalle Cliente", "Sucursal", "Proyectos", "Area",
                                        "Centro de Costos", "Valor Solicitado", "Descripción", "Observaciones", "Observaciones de Anticipo",
                                        "Observaciones de la Solicitud"
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
                                        if (counter == 1)
                                        {
                                            selenium.Click("//span[contains(@id,'ctl00_lblTitulo')]");
                                            Thread.Sleep(100);
                                        }
                                        if (counter == 15) for (int i = 0; i < 14; i++) SendKeys.Send("{DOWN}");
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
                                    for (int i = 0; i < 14; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[1].Click();
                                        Thread.Sleep(500);

                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            SendKeys.Send("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }


                                    //selenium.Click("//a[contains(@id,'btnActualizar')]");
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
        public void SmartPeople_frmNmCtpreNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2_V2.SmartPeople_NTC_4.SmartPeople_frmNmCtpreNTC")
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
                                "HomeEsp", "RetornarEsp", "IdentEsp", "NombreEsp", "NumContEsp", "FechaIniEsp", "FechaFinEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaIni_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaFin_lblFecha']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Identificación", "Nombre", "Nro. de Contrato", "Fecha Inicial",
                                        "Fecha Final"
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


                                    //selenium.Click("//a[contains(@id,'btnActualizar')]");
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
        public void SmartPeople_frmNmCtperLhNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmNmCtperLhNTC")
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
                                "HomeEsp", "CedEsp", "NomEsp", "ApellEsp"
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
                                            Thread.Sleep(2000);

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAPE_EMPL']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Cedula", "Nombre", "Apellido"
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


                                    //selenium.Click("//a[contains(@id,'btnActualizar')]");
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
        public void SmartPeople_frmFdDplinCNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmFdDplinCNTC")
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
                                "HomeEsp", "RetornarEsp", "ParamEsp", "SecuEsp", "CursoEsp", "FechaIniEsp", "FechaFinEsp",
                                "CuposEsp", "NumParInEsp", "SesEsp", "FechaIni2Esp", "FechaFin2Esp", "HoraIniEsp", "HorafinEsp",
                                "LugSesEsp", "ContCurEsp", "DigIdentEsp"
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

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(CamposTotalesMTM[3]);
                                    Thread.Sleep(2000);

                                    ////a[contains(text(),'Detalle')]

                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    //Error 
                                    selenium.Click("//div[4]/div/button");
                                    Thread.Sleep(500);


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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label1']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label3']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecinic_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecfina_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label5']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label6']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecinicS_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecfinaS_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label8']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label9']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label10']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label11']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label12']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Parametro", "Secuencial", "Curso", "Fecha inicial", "Fecha final",
                                        "Cupos", "Nro. particiapntes inscritos", "Sesión", "Fecha inicial",
                                        "Fecha Final", "Hora inicial", "Hora final", "Lugar de la sesión",
                                        "Contenido del curso", "Digite identificación"
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
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }

                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                        if (counter == 12) for (int i = 0; i < 6; i++) SendKeys.Send("{DOWN}");
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
                                    for (int i = 0; i < 6; i++) SendKeys.Send("{UP}");
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


                                    //selenium.Click("//a[contains(@id,'btnActualizar')]");
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
        public void SmartPeople_frmFdNeforDNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmFdNeforDNTC")
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
                                "HomeEsp", "RetornarEsp", "NumContEsp", "IdentEsp", "CodInterEsp", "NameEsp", "ApellEsp",
                                "CentCostEsp", "CargoEsp", "FechaInsEsp", "ConsEsp", "EstSolEsp", "EstBreEsp", "RegisEsp", "RequeEsp",
                                "EspeEsp", "CursoEsp", "OtCurEsp", "PerspecEsp", "ObjAsoEsp", "JustEsp", "ObservEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblApeEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCcos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecInsc']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRmtGene']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstSoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstAdos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodRegi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodRequ']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEspe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCurs']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomCurs']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodPers']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodObes']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_txtJusSoli_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_txtObsErva_lblTexto']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Nro. Contrato", "Identificación", "Cod. Interno", "Nombre", "Apellido",
                                        "Centro de Costo", "Cargo", "Fecha inscripción", "Consecutivo", "Estado de solicitud",
                                        "Estado Brecha", "Registro", "Requerimiento", "Especificación", "Curso", "Otro curso",
                                        "Perspectivas", "Objetivos asociados", "Justificación", "Observaciones"
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
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }

                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                        if (counter == 14) for (int i = 0; i < 11; i++) SendKeys.Send("{DOWN}");
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
                                    for (int i = 0; i < 11; i++) SendKeys.Send("{UP}");
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


                                    //selenium.Click("//a[contains(@id,'btnActualizar')]");
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
        public void SmartPeople_frmLINmCtprePNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_4.SmartPeople_frmLINmCtprePNTC")
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
                                "HomeEsp", "RetornarEsp",
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTip_docu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCedula']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNombres']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTip_docu0']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTip_docu1']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTip_docu2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTip_docu3']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno20']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno21']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTip_docu6']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lbljefInm']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomFeje']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno4']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno5']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label6']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno0']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno1']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno6']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno8']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno9']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno12']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno13']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno16']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno14']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno15']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno18']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno19']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno22']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRno23']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Tipo documento", "Identificación", "Apellidos y Nombres", "Cargo",
                                        "Nivel", "Grado", "Rol", "Año", "Mes", "Ubicación", "Identificación Jefe",
                                        "Nombre Jefe", "Cantidad", "Cantidad a Pagar", "Cantidad", "Catidad a Pagar",
                                        "Cantidad", "Cantidad a Pagar", "Cantidad", "Cantidad a Pagar", ""
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
                                            campoName = selenium.CamposEmergentes(element.Key, element.Value, file);
                                        }

                                        CamposPagina.Add(campoName);
                                        counter++;
                                        Thread.Sleep(100);
                                        if (counter == 14) for (int i = 0; i < 11; i++) SendKeys.Send("{DOWN}");
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
                                    for (int i = 0; i < 11; i++) SendKeys.Send("{UP}");
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


                                    //selenium.Click("//a[contains(@id,'btnActualizar')]");
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
