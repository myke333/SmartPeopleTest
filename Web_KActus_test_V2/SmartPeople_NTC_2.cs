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
using System.Linq;


namespace Web_Kactus_Test_V2
{
    /// <summary>
    /// Descripción resumida de SmartPeople_NTC_2
    /// </summary>
    [TestClass]
    public class SmartPeople_NTC_2 : FuncionesVitales
    {
        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();
        public SmartPeople_NTC_2()
        {
        }
       
        [TestMethod]
        public void PruebaGrabacion()
        {
            //System.Diagnostics.Debugger.Launch();
            // public void SmartPeople_frmBIEdforNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.PruebaGrabacion")
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

                            if (
                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                rows["AgregarEsp"].ToString().Length != 0 && rows["AgregarEsp"].ToString() != null &&
                                rows["CodCargEsp"].ToString().Length != 0 && rows["CodCargEsp"].ToString() != null &&
                                rows["NomCargEsp"].ToString().Length != 0 && rows["NomCargEsp"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string url2 = rows["url2"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                string AgregarEsp = rows["AgregarEsp"].ToString();
                                string CodCargEsp = rows["CodCargEsp"].ToString();
                                string NomCargEsp = rows["NomCargEsp"].ToString();

                                try
                                {
                                    string database = "";
                                    if (url.ToLower().Contains("ora"))
                                    {
                                        database = "ORA";
                                    }
                                    else
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);


                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);

                                    selenium.Screenshot("Manual de funciones", true, file);


                                    // Validacion titulo
                                    string Titulo = selenium.Title();
                                    Thread.Sleep(100);
                                    if (Titulo != TituloEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El título es incorrecto, el esperado es: " + TituloEsp + " y el encontrado es: " + Titulo);
                                    }
                                    Thread.Sleep(100);

                                    //Validación subtitulo botones
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    Thread.Sleep(100);
                                    if (Subtitulo != SubtituloEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El subtítulo es incorrecto, el esperado es: " + SubtituloEsp + " y el encontrado es: " + Subtitulo);
                                    }
                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);

                                    //boton agregar
                                    //aqui escribe su codigo
                                    string Agregar = selenium.EmergenteBotones("ctl00_btnNuevo");
                                    Thread.Sleep(500);
                                    if (Agregar != AgregarEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + AgregarEsp + " y el encontrado es: " + Agregar);
                                    }


                                    //Validación Emergentes campos
                                    string CodCarg = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblConsuCodCarg')]", "Consultar Codigo Cargo", file);
                                    if (CodCarg != CodCargEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + CodCargEsp + " y el encontrado es: " + CodCarg);
                                    }
                                    Thread.Sleep(100);

                                    string NomCarg = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblConsuNomCargo')]", "Consultar Nombre del Cargo", file);
                                    if (NomCarg != NomCargEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + NomCargEsp + " y el encontrado es: " + NomCarg);
                                    }
                                    Thread.Sleep(100);

                                    /////////// Validación TABS ////////
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$txt')]")));
                                    elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
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
        public void SmartPeople_frmAcMafunDNTC()
        {
            // public void SmartPeople_frmBIEdforNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmAcMafunDNTC")
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

                            if (
                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                rows["RetornarEsp"].ToString().Length != 0 && rows["RetornarEsp"].ToString() != null &&
                                rows["CodCargEsp"].ToString().Length != 0 && rows["CodCargEsp"].ToString() != null &&
                                rows["NomCargEsp"].ToString().Length != 0 && rows["NomCargEsp"].ToString() != null &&
                                rows["VersionEsp"].ToString().Length != 0 && rows["VersionEsp"].ToString() != null &&
                                rows["VigenciaDesdeEsp"].ToString().Length != 0 && rows["VigenciaDesdeEsp"].ToString() != null &&
                                rows["FechaElabEsp"].ToString().Length != 0 && rows["FechaElabEsp"].ToString() != null &&
                                rows["CodGrupoEsp"].ToString().Length != 0 && rows["CodGrupoEsp"].ToString() != null &&
                                rows["CodigoEsp"].ToString().Length != 0 && rows["CodigoEsp"].ToString() != null &&
                                rows["EstadoEsp"].ToString().Length != 0 && rows["EstadoEsp"].ToString() != null &&
                                rows["ElaboradoEsp"].ToString().Length != 0 && rows["ElaboradoEsp"].ToString() != null &&
                                rows["AprobadoEsp"].ToString().Length != 0 && rows["AprobadoEsp"].ToString() != null &&
                                rows["ValidadoEsp"].ToString().Length != 0 && rows["ValidadoEsp"].ToString() != null &&
                                rows["AutorizadoEsp"].ToString().Length != 0 && rows["AutorizadoEsp"].ToString() != null &&
                                rows["DesCambEsp"].ToString().Length != 0 && rows["DesCambEsp"].ToString() != null &&
                                rows["ObserSoliEsp"].ToString().Length != 0 && rows["ObserSoliEsp"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string url2 = rows["url2"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                string RetornarEsp = rows["RetornarEsp"].ToString();
                                string CodCargEsp = rows["CodCargEsp"].ToString();
                                string NomCargEsp = rows["NomCargEsp"].ToString();
                                string VersionEsp = rows["VersionEsp"].ToString();
                                string VigenciaDesdeEsp = rows["VigenciaDesdeEsp"].ToString();
                                string FechaElabEsp = rows["FechaElabEsp"].ToString();
                                string CodGrupoEsp = rows["CodGrupoEsp"].ToString();
                                string CodigoEsp = rows["CodigoEsp"].ToString();
                                string EstadoEsp = rows["EstadoEsp"].ToString();
                                string ElaboradoEsp = rows["ElaboradoEsp"].ToString();
                                string AprobadoEsp = rows["AprobadoEsp"].ToString();
                                string ValidadoEsp = rows["ValidadoEsp"].ToString();
                                string AutorizadoEsp = rows["AutorizadoEsp"].ToString();
                                string DesCambEsp = rows["DesCambEsp"].ToString();
                                string ObserSoliEsp = rows["ObserSoliEsp"].ToString();
                                try
                                {
                                    string database = "";
                                    if (url.ToLower().Contains("ora"))
                                    {
                                        database = "ORA";
                                    }
                                    else
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);


                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);

                                    ////a[contains(text(),'Detalle')]
                                    selenium.Click("//a[@id='ctl00_btnCerrar']");
                                    Thread.Sleep(6000);


                                    selenium.Screenshot("Manual de funciones", true, file);


                                    // Validacion titulo
                                    string Titulo = selenium.Title();
                                    Thread.Sleep(100);
                                    if (Titulo != TituloEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El título es incorrecto, el esperado es: " + TituloEsp + " y el encontrado es: " + Titulo);
                                    }
                                    Thread.Sleep(100);

                                    //Validación subtitulo botones
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    Thread.Sleep(100);
                                    if (Subtitulo != SubtituloEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El subtítulo es incorrecto, el esperado es: " + SubtituloEsp + " y el encontrado es: " + Subtitulo);
                                    }
                                    Thread.Sleep(100);


                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);

                                    //boton agregar
                                    //aqui escribe su codigo
                                    string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");
                                    Thread.Sleep(500);
                                    if (Retornar != RetornarEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + RetornarEsp + " y el encontrado es: " + Retornar);
                                    }

                                    //Validación Emergentes campos


                                    string CodCarg = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblCodCargo')]", "Consultar Codigo Cargo", file);
                                    string NomCarg = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblNomCargo')]", "Consultar Nombre del Cargo", file);
                                    string Version = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblVerSion')]", "Consultar Versión", file);
                                    string VigenciaDesde = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblVigdesd')]", "Consultar Vigencia Desde", file);
                                    string FechaElaboracion = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_KCtrlFecCrea_lblFecha')]", "Consultar Fecha Elaboración", file);
                                    string CodGrupo = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblCodGrup')]", "Consultar Grupo", file);
                                    string Codigo = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblVCodForm')]", "Consultar Codigo", file);
                                    string Estado = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblEstVers')]", "Consultar Estado", file);
                                    string Elaborado = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblElaBora')]", "Consultar Elaborado", file);
                                    string Aprobado = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblAprObad')]", "Consultar Aprobado", file);
                                    string Validado = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblCodVali')]", "Consultar Validado", file);
                                    string Autorizado = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblCodAuto')]", "Consultar Autorizado", file);
                                    string DesCamb = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_lblDesCamb')]", "Comsultar Descripcion del Cambio", file);
                                    string ObserSoli = selenium.CamposEmergentes("//span[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_lblTexto')]", "Consultar Observaciones de la Solicitud", file);


                                    List<string> CamposPagina = new List<string>() { CodCarg, NomCarg, Version, VigenciaDesde, FechaElaboracion, CodGrupo, Codigo,
                                                                                    Estado, Elaborado, Aprobado, Validado, Autorizado, DesCamb, ObserSoli };

                                    List<string> CamposMTM = new List<string>() { CodCargEsp, NomCargEsp, VersionEsp, VigenciaDesdeEsp, FechaElabEsp, CodGrupoEsp, CodigoEsp,
                                                                                 EstadoEsp, ElaboradoEsp, AprobadoEsp, ValidadoEsp, AutorizadoEsp, DesCambEsp, ObserSoliEsp };


                                    var elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);


                                    foreach (var campo in elementos)
                                    {
                                        if (campo.Key != campo.Value)
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                        }
                                        Thread.Sleep(100);
                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$txt')]")));
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
        public void SmartPeople_frmBiEdforNTC()
        {
            // public void SmartPeople_frmBIEdforNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmBiEdforNTC")
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

                            if (
                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                rows["RetornarEsp"].ToString().Length != 0 && rows["RetornarEsp"].ToString() != null &&
                                rows["GuardarEsp"].ToString().Length != 0 && rows["GuardarEsp"].ToString() != null &&
                                rows["ModalidadEsp"].ToString().Length != 0 && rows["ModalidadEsp"].ToString() != null &&
                                rows["NomEstudiosEsp"].ToString().Length != 0 && rows["NomEstudiosEsp"].ToString() != null &&
                                rows["NomEspeEsp"].ToString().Length != 0 && rows["NomEspeEsp"].ToString() != null &&
                                rows["InstitucionEsp"].ToString().Length != 0 && rows["InstitucionEsp"].ToString() != null &&
                                rows["CiudadEsp"].ToString().Length != 0 && rows["CiudadEsp"].ToString() != null &&
                                rows["MetodologiaEsp"].ToString().Length != 0 && rows["MetodologiaEsp"].ToString() != null &&
                                rows["TituloConEsp"].ToString().Length != 0 && rows["TituloConEsp"].ToString() != null &&
                                rows["FechaInicioEsp"].ToString().Length != 0 && rows["FechaInicioEsp"].ToString() != null &&
                                rows["FechaFinalEsp"].ToString().Length != 0 && rows["FechaFinalEsp"].ToString() != null &&
                                rows["TiempoEsEsp"].ToString().Length != 0 && rows["TiempoEsEsp"].ToString() != null &&
                                rows["TerminadoEsp"].ToString().Length != 0 && rows["TerminadoEsp"].ToString() != null &&
                                rows["EstudiaAcEsp"].ToString().Length != 0 && rows["EstudiaAcEsp"].ToString() != null &&
                                rows["EstudioInterEsp"].ToString().Length != 0 && rows["EstudioInterEsp"].ToString() != null &&
                                rows["GraduadoEsp"].ToString().Length != 0 && rows["GraduadoEsp"].ToString() != null &&
                                rows["FechaGradoEsp"].ToString().Length != 0 && rows["FechaGradoEsp"].ToString() != null &&
                                rows["PromedioEsp"].ToString().Length != 0 && rows["PromedioEsp"].ToString() != null &&
                                rows["DatosVerEsp"].ToString().Length != 0 && rows["DatosVerEsp"].ToString() != null &&
                                rows["MatProEsp"].ToString().Length != 0 && rows["MatProEsp"].ToString() != null &&
                                rows["FechaExpEsp"].ToString().Length != 0 && rows["FechaExpEsp"].ToString() != null &&
                                rows["TramiteEsp"].ToString().Length != 0 && rows["TramiteEsp"].ToString() != null &&
                                rows["CertDiploEsp"].ToString().Length != 0 && rows["CertDiploEsp"].ToString() != null &&
                                rows["TipoDocEsp"].ToString().Length != 0 && rows["TipoDocEsp"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string url2 = rows["url2"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                string RetornarEsp = rows["RetornarEsp"].ToString();
                                string GuardarEsp = rows["GuardarEsp"].ToString();
                                string ModalidadEsp = rows["ModalidadEsp"].ToString();
                                string NomEstudiosEsp = rows["NomEstudiosEsp"].ToString();
                                string NomEspeEsp = rows["NomEspeEsp"].ToString();
                                string InstitucionEsp = rows["InstitucionEsp"].ToString();
                                string CiudadEsp = rows["CiudadEsp"].ToString();
                                string MetodologiaEsp = rows["MetodologiaEsp"].ToString();
                                string TituloConEsp = rows["TituloConEsp"].ToString();
                                string FechaInicioEsp = rows["FechaInicioEsp"].ToString();
                                string FechaFinalEsp = rows["FechaFinalEsp"].ToString();
                                string TiempoEsEsp = rows["TiempoEsEsp"].ToString();
                                string TerminadoEsp = rows["TerminadoEsp"].ToString();
                                string EstudiaAcEsp = rows["EstudiaAcEsp"].ToString();
                                string EstudioInterEsp = rows["EstudioInterEsp"].ToString();
                                string GraduadoEsp = rows["GraduadoEsp"].ToString();
                                string FechaGradoEsp = rows["FechaGradoEsp"].ToString();
                                string PromedioEsp = rows["PromedioEsp"].ToString();
                                string DatosVerEsp = rows["DatosVerEsp"].ToString();
                                string MatProEsp = rows["MatProEsp"].ToString();
                                string FechaExpEsp = rows["FechaExpEsp"].ToString();
                                string TramiteEsp = rows["TramiteEsp"].ToString();
                                string CertDiploEsp = rows["CertDiploEsp"].ToString();
                                string TipoDocEsp = rows["TipoDocEsp"].ToString();

                                try
                                {
                                    string database = "";
                                    if (url.ToLower().Contains("ora"))
                                    {
                                        database = "ORA";
                                    }
                                    else
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);


                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    ChromeDriver driver = selenium.returnDriver();
                                    driver.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);

                                    ////a[contains(text(),'Detalle')]
                                    //selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                    //Thread.Sleep(500);


                                    selenium.Screenshot("Mi Educ. Formal", true, file);


                                    // Validacion titulo
                                    string Titulo = selenium.Title();
                                    Thread.Sleep(100);
                                    if (Titulo != TituloEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El título es incorrecto, el esperado es: " + TituloEsp + " y el encontrado es: " + Titulo);
                                    }
                                    Thread.Sleep(100);

                                    //Validación subtitulo botones
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    Thread.Sleep(100);
                                    if (Subtitulo != SubtituloEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El subtítulo es incorrecto, el esperado es: " + SubtituloEsp + " y el encontrado es: " + Subtitulo);
                                    }
                                    Thread.Sleep(100);


                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);

                                    //boton agregar
                                    //aqui escribe su codigo
                                    string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");
                                    Thread.Sleep(500);
                                    if (Retornar != RetornarEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + RetornarEsp + " y el encontrado es: " + Retornar);
                                    }

                                    //boton Guardar
                                    //aqui escribe su codigo
                                    string Guardar = selenium.EmergenteBotones("btnGuardar");
                                    Thread.Sleep(500);
                                    if (Guardar != GuardarEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Guardar es incorrecto, el esperado es: " + GuardarEsp + " y el encontrado es: " + Guardar);
                                    }

                                    //Lista de Variables de MTM
                                    List<string> CamposMTM = new List<string>() { ModalidadEsp, NomEstudiosEsp, NomEspeEsp, InstitucionEsp, CiudadEsp, MetodologiaEsp, TituloConEsp,
                                                                                 FechaInicioEsp, FechaFinalEsp, TiempoEsEsp, TerminadoEsp, EstudiaAcEsp, EstudioInterEsp, GraduadoEsp,
                                                                                 FechaGradoEsp, PromedioEsp, MatProEsp, FechaExpEsp, CertDiploEsp, TipoDocEsp};



                                    //Validación Emergentes campos
                                    var campos = new Dictionary<string, string>
                                    {
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblNomModi')]"] = "Modalidad",
                                        ["//div[@id='printable']/div[2]/div[2]/div/div/span"] = "Nombre de los Estudios",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_KCtrlTxtValMNomEspe_lblTexto')]"] = "Nombre Específico",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblNomInst')]"] = "Institución",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_lblDivPoli')]"] = "Ciudad",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblCodItem')]"] = "Metodologia",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblTitConv')]"] = "Titulo convalidado ante el MEN MINISTERIO DE EDUCACION NACIONAL",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_kcfFecInic_lblFecha')]"] = "Fecha Inicio",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_kcfFecTerm_lblFecha')]"] = "Fecha Final",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblUniTiem')]"] = "Tiempo de Estudio",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblEstTerm')]"] = "Terminado",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblEstActu')]"] = "Estudia Actualmente",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblEstInte')]"] = "Estudio Interrumpido",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblGraDuad')]"] = "Graduado",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_kcfFecGrad_lblFecha')]"] = "Fecha de Grado",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblProCarr')]"] = "Promedio",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblMatProf')]"] = "Matrícula Profesional",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_kcfFecExtp_lblFecha')]"] = "Fecha Expedición",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblAdjunto')]"] = "Certificaciones y Diplomas",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_KCtrTipoDocumento_lblTIP_DOCU')]"] = "Tipo de documento"


                                    };

                                    List<string> CamposPagina = new List<string>();
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
                                        if (counter == 4) for (int i = 0; i < 8; i++) SendKeys.Send("{DOWN}");
                                    }

                                    //Diccionario con CamposEsperados y CamposObtenidos
                                    elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    //Verificación de errores en los campos
                                    foreach (var campo in elementos)
                                    {

                                        if (campo.Key != campo.Value)
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                        }

                                    }


                                    /////////// Validación TABS ////////
                                    for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
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
        public void SmartPeople_frmBiEdnfoNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmBiEdnfoNTC")
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

                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp", "RetornarEsp", "GuardarEsp", "ModalidadEsp", "NomEstudiosEsp", "NomEspeEsp",
                                "InstitucionEsp", "FechaInicioEsp", "FechaFinalEsp", "CiudadEsp",  "TerminadoEsp", "EstudiaAcEsp",
                                "TiempoEsEsp", "CertDiploEsp", "TipoDocEsp"
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
                                    //selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                    //Thread.Sleep(500);


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
                                        if (nameControl.Key != nameControl.Value)
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de" + name[controlCounter] + " incorrecto, el esperado es: " + nameControl.Value + " y el encontrado es: " + nameControl.Key);
                                        }
                                        controlCounter++;
                                        Thread.Sleep(100);
                                    }


                                    //Validación Emergentes campos
                                    var campos = new Dictionary<string, string>
                                    {
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblModAcad')]"] = "Modalidad",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblNomEstu')]"] = "Nombre de los Estudios",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_KCtrlTxtValMNomEspe_lblTexto')]"] = "Nombre Específico",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblNomInst')]"] = "Institución",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_kcfFecInic_lblFecha')]"] = "Fecha Inicio",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_kcfFecTerm_lblFecha')]"] = "Fecha Final",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_lblDivPoli')]"] = "Ciudad",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblEstTerm')]"] = "Terminado",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblEstActu')]"] = "Estudia Actualmente",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblUniTiem')]"] = "Tiempo de Estudio",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblAdjunto')]"] = "Certificaciones o Diplomas",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_KCtrTipoDocumento_lblTIP_DOCU')]"] = "Tipo de documento"

                                    };

                                    List<string> CamposPagina = new List<string>();
                                    List<string> CamposMTM = CamposTotalesMTM.GetRange(9, (CamposTotalesMTM.Count - 9));
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
                                        if (counter == 4) for (int i = 0; i < 8; i++) SendKeys.Send("{DOWN}");
                                    }


                                    elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    foreach (var campo in elementos)
                                    {

                                        if (campo.Key != campo.Value)
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                        }

                                    }


                                    /////////// Validación TABS ////////
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
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
        public void SmartPeople_frmBiEmpdoNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmBiEmpdoNTC")
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

                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp", "RetornarEsp", "GuardarEsp", "IdentificacionEsp", "NombresEsp",
                                "ApellidosEsp", "CodigoEsp", "NumeroEsp", "CiudadEsp", "FechaExpEsp",
                                "FechaVenEsp", "ObserEsp", "CertDiploEsp"
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
                                    //selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                    //Thread.Sleep(500);


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
                                        if (nameControl.Key != nameControl.Value)
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de" + name[controlCounter] + " incorrecto, el esperado es: " + nameControl.Value + " y el encontrado es: " + nameControl.Key);
                                        }
                                        controlCounter++;
                                        Thread.Sleep(100);
                                    }


                                    //Validación Emergentes campos
                                    var campos = new Dictionary<string, string>
                                    {
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblCodEmpl')]"] = "Identificación",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblNomEmpl')]"] = "Nombres",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblApeEmpl')]"] = "Apellidos",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblCodDocu')]"] = "Código",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblNumDocu')]"] = "Número",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_KCtrDivipExpe_lblDivPoli')]"] = "Ciudad",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_kcfFecExpe_lblFecha')]"] = "Fecha Expedición",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_kcfFecVenci_lblFecha')]"] = "Fecha Vencimiento",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_KCtrlTxtValObser_lblTexto')]"] = "Observaciones",
                                        ["//span[contains(@id,'ctl00_ContenidoPagina_lblAdjunto')]"] = "Certificaciones o Diplomas"

                                    };

                                    List<string> CamposPagina = new List<string>();
                                    List<string> CamposMTM = CamposTotalesMTM.GetRange(9, (CamposTotalesMTM.Count - 9));
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
                                        if (counter == 5) for (int i = 0; i < 6; i++) SendKeys.Send("{DOWN}");
                                    }


                                    elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    foreach (var campo in elementos)
                                    {

                                        if (campo.Key != campo.Value)
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                        }

                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 10; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
                                    elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
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

                                    //selenium.Click("//a[contains(@id,'ctl00_btnNuevo')]");
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(500);
                                    //xpath=//a[@id='btnGuardar']
                                    selenium.Screenshot("Campos Necesarios", true, file);


                                    for (int k = 0; k < 10; k++) SendKeys.Send("{UP}");
                                    Thread.Sleep(100);
                                    selenium.Screenshot("Campos Necesarios", true, file);


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
        public void SmartPeople_frmBiEdnfoCNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmBiEdnfoCNTC")
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
                            int indexVariables = 7;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp", "ModalidadEsp", "FechaInicioEsp", "FechaFinalEsp", "TiempoEsEsp",
                                "CiudadEsp", "ConsulCedEsp", "ConsulNomEsp", "ConsulApellEsp", "ResponsableEsp"
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
                                    //selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                    //Thread.Sleep(500);


                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6]

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


                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblModAcad']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_kcfFecInic_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_kcfFecTerm_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblUniTiem']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrDivipUbi_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrLupaEmpl1_lblConsuCedEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrLupaEmpl1_lblConsuNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrLupaEmpl1_lblConsuAplEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrLupaEmpl1_lblCodEmpl']"
                                    };

                                    for (int i = 0; i < xpath.Count; i++)
                                    {
                                        xpath[i] = xpath[i].Replace("xpath=", "");
                                        xpath[i] = xpath[i].Replace("=", ",");
                                        xpath[i] = xpath[i].Insert(7, "contains(");
                                        xpath[i] = xpath[i].Insert(xpath[i].Length - 1, ")");
                                    }

                                    //Validación Emergentes campos
                                    var campos = new Dictionary<string, string>
                                    {
                                        [xpath[0]] = "Modalidad",
                                        [xpath[1]] = "Fecha Inicio",
                                        [xpath[2]] = "Fecha Final",
                                        [xpath[3]] = "Tiempo de Estudio",
                                        [xpath[4]] = "Ciudad",
                                        [xpath[5]] = "Consultar Por Cedula",
                                        [xpath[6]] = "Consultar Por Nombres",
                                        [xpath[7]] = "Consultar Por Apellidos",
                                        [xpath[8]] = "Responsables"

                                    };

                                    List<string> CamposPagina = new List<string>();
                                    List<string> CamposMTM = CamposTotalesMTM.GetRange(indexVariables, (CamposTotalesMTM.Count - indexVariables));
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
                                        if (counter == 5) for (int i = 0; i < 6; i++) SendKeys.Send("{DOWN}");
                                    }


                                    elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    foreach (var campo in elementos)
                                    {

                                        if (campo.Key != campo.Value)
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                        }

                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 10; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
                                    elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
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
        public void SmartPeople_frmBiFamilNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmBiFamilNTC")
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
                            int indexVariables = 9;
                            List<string> CamposTotalesMTM = new List<string>();
                            List<string> variablesMTM = new List<string>() {
                                "EmpleadoUser", "EmpleadoPass", "url", "url2", "TituloEsp", "SubtituloEsp",
                                "HomeEsp", "RetornarEsp", "GuardarEsp", "IdentificacionEsp",
                                "TipoDocEsp", "RegCivEsp", "NombresEsp", "SegNomEsp", "TipoRelaEsp", "ApellidosEsp", "SegApellEsp",
                                "SexoEsp", "FechaNacEsp", "ViveEsp", "BenefiEsp", "GrupSangEsp", "FactSangEsp", "EstadoCivilEsp",
                                "FechaMatConEsp", "CiudadEsp", "DireccionEsp", "TelefonoEsp", "ActividadEsp", "EstablecEsp",
                                "EscolaridadEsp", "HobbiesEsp", "TrabaOtraEntEsp", "Discapacidad", "BenefiCajaConEsp",
                                "EsDepEsp", "FechaRatEsp", "EsBenEsp", "DocumentosEsp", "TipoDeDocEsp"
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
                                    //selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                    //Thread.Sleep(500);


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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodFami']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipIden']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodRcvl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomFami1']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomFami2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipRela']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblApeFami1']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblApeFami2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSexFami']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_kcfFecNaci_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstVida']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorBene']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblGruSang']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFacSang']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstCivi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KcfFecMatr_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrDivip_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDirFami']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTelFami']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTraEstu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSitEstu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblGraEsco']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblHobFami']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblOtrEnti']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstDisc']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblBenCaco']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFamDepe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecVncr_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblBenEps']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAdjunto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrTipoDocumento_lblTIP_DOCU']"

                                    };

                                    List<string> descripcionSC = new List<string>() {
                                        "Identificación", "Tipo Documento", "Registro Civil o NUIP o Pasaporte", "Nombres", "Segundo Nombre",
                                        "Tipo de Relacion", "Apellidos", "Segundo Apellido", "Sexo", "Fecha de Nacimiento", "Vive",
                                        "Beneficiario", "Grupo de Sangre", "Factor Sanguíneo", "Estado Civil", "Fecha Matrimonio o Convivencia",
                                        "Ciudad", "Dirección", "Teléfono", "Actividad", "Establecimiento", "Escolaridad", "Hobbies", "Trabaja con Otra Entidad",
                                        "Discapacidad", "Beneficiario a Caja de Compensación", "Es Dependiende del Empleado", "Fecha Ratificación", "Es Beneficiario de EPS",
                                        "Documentos", "Tipo de documento"
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
                                        if (counter == 12) for (int i = 0; i < 8; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 21) for (int i = 0; i < 8; i++) SendKeys.Send("{DOWN}");
                                    }


                                    elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    foreach (var campo in elementos)
                                    {

                                        if (campo.Key != campo.Value)
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                        }

                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");

                                    List<IWebElement> elementList = new List<IWebElement>();
                                    List<IWebElement> elementListPagina = new List<IWebElement>();
                                    List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);

                                    elementListPagina.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

                                    if (elementListPagina.Count > 0)
                                    {
                                        //int cont = 0;
                                        foreach (IWebElement pageEle in elementListPagina)
                                        {
                                            // cont++;
                                            Thread.Sleep(800);

                                            if (pageEle.TagName == "select" || pageEle.TagName == "input")
                                            {
                                                if (pageEle.Displayed && pageEle.Enabled)
                                                {
                                                    String id = pageEle.GetAttribute("id");
                                                    if (id == "ctl00_ContenidoPagina_tcBeSolic_TabPDatos_txtMunicipioInst" || id == "ctl00_ContenidoPagina_tcBeSolic_TabPDatos_ddlCodMone" || id == "ctl00_ContenidoPagina_tcBeSolic_TabPDatos_txtValTarm" || id == "ctl00_ContenidoPagina_tcBeSolic_TabPDatos_txtValBeca")
                                                    {
                                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El control de ID " + id + " no debe permitir la opcion de tabular");
                                                    }

                                                    SendKeys.Send("{TAB}");
                                                    selenium.Screenshot("TAB", true, file);

                                                    Thread.Sleep(100);

                                                }
                                                else
                                                {
                                                    SendKeys.Send("{TAB}");
                                                    selenium.Screenshot("TAB", true, file);

                                                    Thread.Sleep(100);
                                                }
                                            }

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
        public void SmartPeople_frmBiHvextNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmBiHvextNTC")
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
                                "HomeEsp", "RetornarEsp", "GuardarEsp", "EmpresaEsp", "DireccionEsp", "TelefonoEsp",
                                "TipoEmpreEsp", "CorreoEntEsp", "SalarioDevenEsp", "EmpleoAcEsp", "CargoEjecEsp", "DedicacionEsp",
                                "FechaIngreEsp", "FechaRetiroEsp", "ManejaPerEsp", "CargoDesemEsp", "Area", "ProLogEsp", "TipoConEsp",
                                "MotivRetiroEsp", "HerramientasEsp", "Ciudad", "JefeInmediatoEsp", "CargJefeInmeEsp", "TiempoServEsp",
                                "ActivEmpreEsp", "FuncionesEsp", "AreasExpeEsp", "CertDiploEsp", "TipoDocEsp"
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
                                    //selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                    //Thread.Sleep(500);

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDirEmpr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTelEmpr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipEmpr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEntMail']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSalDemp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEmpActu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCarEjec']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDedIcac']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_kcfFecIngr_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_kcfFecReti_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblManPers']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCarDese']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDepEmpr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblProLogr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMotReti']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRlHerra']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrDivipEmpre_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblJefInme']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCarJefe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTieServ']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblActEmp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtValMFunReal_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAreExp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAdjunto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrTipoDocumento_lblTIP_DOCU']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Empresa", "Dirección", "Teléfono", "Tipo de Empresa", "Correo de la Entidad", "Salario Devengado", "Empleo Actual",
                                        "Cargo Ejecutivo", "Dedicación", "Fecha de Ingreso", "Fecha de Retiro", "Maneja personal", "Cargo Desempeñado",
                                        "Area", "Proyectos Y Logros", "Tipo de COntrato", "Motivo de Retiro", "Herramientas", "Ciudad", "Jefe Inmediato",
                                        "Cargo Jefe Inmediato", "Tiempo de Servicio", "Actividad Empresa", "Funciones",
                                        "Areas de Experiencia", "Certificaciones o Diplomas", "Tipo de documento"
                                    };
                                    // Debugger.Launch();
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
                                        if (counter == 19) for (int i = 0; i < 13; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 25) for (int i = 0; i < 3; i++) SendKeys.Send("{DOWN}");
                                    }

                                    Console.WriteLine(CamposPagina.Count);
                                    elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    foreach (var campo in elementos)
                                    {

                                        if (campo.Key != campo.Value)
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                        }

                                    }


                                    /////////// Validación TABS ////////
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
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
        public void SmartPeople_frmBiPeFunRNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmBiPeFunRNTC")
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
                                "HomeEsp", "RetornarEsp", "CodRolEsp", "NomRolEsp", "CargoEsp", "NivelEsp", "GradCarEsp",
                                "CodFormatEsp", "VigDesEsp", "VigHastaEsp", "VersionEsp", "ElabPorEsp", "AprobPorEsp", "TiposVinRolEsp",
                                "ManualEsp", "IndiAcEsp", "NumResEsp", "FechaResEsp"
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
                                    //selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                    //Thread.Sleep(500);

                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Retornar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Retornar = selenium.EmergenteBotones("ctl00_btnRetornar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        [Retornar] = CamposTotalesMTM[7]

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodRol']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomRol']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodNivel']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblGraCarg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodForm']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecCrea_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlVigHast_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblVerSion']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblElaBora']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAprObad']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblGrpEval']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodManu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblIndActi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroReso']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecReso_lblFecha']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Código de Rol", "Nombre de Rol", "Cargo", "Nivel", "Grado del Cargo", "Código de Formato",
                                        "Vigencia Desde", "Vigencia Hasta", "Versión", "Elaborado Por", "Aprobado Por", "Tipos de Vinculación de Rol",
                                        "Manual", "Indicador de Actividad", "Número de Resolución", "Fecha de Resolución"
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

                                    Console.WriteLine(CamposPagina.Count);
                                    elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    foreach (var campo in elementos)
                                    {

                                        if (campo.Key != campo.Value)
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                        }

                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    //for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
                                    for (int i = 0; i < 3; i++) SendKeys.Send("{DOWN}");
                                    //selenium.Click("//a[contains(@id,'ctl00_btnNuevo')]");
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnGuardRol')]");
                                    Thread.Sleep(500);
                                    //xpath=//a[@id='btnGuardar']
                                    selenium.Screenshot("Campos Necesarios", true, file);

                                    //Relative: xpath=//div[@id='printable']/div[17]/div[3]
                                    //Position: xpath=//div[17]/div[3]



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
        public void SmartPeople_frmBpSopreNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmBpSopreNTC")
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
                                "HomeEsp", "GuardarEsp", "NumContratoEsp", "IdentEsp", "NomApelEsp", "CodInterEsp", "CargoEsp",
                                "FechaInEsp", "TipoSalaEsp", "SueldoBaEsp", "IngresosEsp", "EstCivilEsp", "IngreFamEsp", "DedExternEsp",
                                "CompViEsp", "FechaSolCorEsp", "NumRadEsp", "NumeroEsp", "DescPresEsp", "ValInmuEsp", "ValRemEsp",
                                "PlazoAñosEsp", "ValSolEsp", "EstSolEsp"
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
                                    //selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                    //Thread.Sleep(500);

                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Guardar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        [Guardar] = CamposTotalesMTM[7]

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNmbEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecIngr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipSala']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSueBasi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTotInmt']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstCivi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblIngFami']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDedExte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblComVige']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_kcfFecSoliA_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroRadi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRmtSoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodPres']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValInmu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValRemo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPlaPres']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValPres']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstSoli']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Numero Contrato", "Identificación", "Nombres y Apellidos", "Codigo Interno", "Cargo",
                                        "Fecha Ingreso", "Tipo de Salario", "Sueldo Básico", "Ingresos", "Estado Civil", "Ingresos Familiares",
                                        "Deducciones Externas", "Compromisos Vigentes", "Fecha Solicitud y Corte", "Número Radicación", "Número",
                                        "Descripción Préstamo", "Valor del Inmueble", "Valor de la Remodelación", "Plazo en Años",
                                        "Valor Solicitado", "Estado de la Solicitud"
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
                                        if (counter == 13) for (int i = 0; i < 12; i++) SendKeys.Send("{DOWN}");

                                    }

                                    Console.WriteLine(CamposPagina.Count);
                                    elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    foreach (var campo in elementos)
                                    {

                                        if (campo.Key != campo.Value)
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                        }

                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 12; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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

                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(500);
                                    //xpath=//a[@id='btnGuardar']
                                    selenium.Screenshot("Campos Necesarios", true, file);

                                    //Relative: xpath=//div[@id='printable']/div[17]/div[3]
                                    //Position: xpath=//div[17]/div[3]



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
        public void SmartPeople_frmBpSopreSiNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmBpSopreSiNTC")
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
                                "HomeEsp", "GuardarEsp", "NumContratoEsp", "IdentEsp", "NomApelEsp", "CodInterEsp",
                                "CargoEsp", "FechaInEsp", "TipoSalaEsp", "SuelBaEsp", "IngresosEsp", "IngreFamEsp", "DedExtEsp",
                                "ComVigEsp", "FechaSolCorEsp", "NumRadEsp", "NumeroEsp", "DescPresEsp", "ValInEsp", "ValRemEsp",
                                "PlazoAñosEsp", "ValSolEsp", "NumCuotasEsp", "PorPagoMenEsp", "ValCuotMenEsp", "PorPagoPriEsp",
                                "ValCoutPriEsp", "EstSolEsp", "TelefEsp", "ObserEsp"
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
                                    ///
                                    try
                                    {
                                        selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                        Thread.Sleep(500);
                                    }
                                    catch (Exception e)
                                    {

                                    }



                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Guardar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        [Guardar] = CamposTotalesMTM[7]

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNmbEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCFFecIngr_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipSala']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSueBasi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTotInmt']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblIngFami']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDedExte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblComVige']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCFFecSoli_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroRadi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRmtSoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodPres']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValInmu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValRemo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPlaPres']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValPres']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNumCuot']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorCume']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValCuot']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPorPexp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValDesc']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstSoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTelEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtObsErva_lblTexto']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Numero de Contrato", "Identificación", "Nombres y Apellidos", "Codigo Interno", "Cargo",
                                        "Fecha Ingreso", "Tipo de Salario", "Sueldo Básico", "Ingresos", "Ingresos Familiares",
                                        "Deducciones Externas", "Compromisos Vigentes", "Fecha Solicitud o Corte", "Numero de Radicación",
                                        "Número", "Descripción Préstamo", "Valor del Inmueble", "Valor de la Remodelación", "Plazo de Años",
                                        "Valor Solicitado", "Numero de Cuotas", "Porcentaje Pago Mensual", "Valor Cuota Mensual", "Porcentaje Pago contra Prima",
                                        "Valor Cuota Prima", "Estado de la Solicitud", "Teléfono", "Observaciones"
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
                                        if (counter == 18) for (int i = 0; i < 12; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 25) for (int i = 0; i < 3; i++) SendKeys.Send("{DOWN}");

                                    }

                                    Console.WriteLine(CamposPagina.Count);
                                    elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    foreach (var campo in elementos)
                                    {

                                        if (campo.Key != campo.Value)
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                        }

                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 16; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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

                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(500);
                                    //xpath=//a[@id='btnGuardar']
                                    selenium.Screenshot("Campos Necesarios", true, file);

                                    //Relative: xpath=//div[@id='printable']/div[17]/div[3]
                                    //Position: xpath=//div[17]/div[3]



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
        public void SmartPeople_frmEdFomeiFNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmEdFomeiFNTC")
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
                                "HomeEsp", "GereOfEsp", "AñoEsp", "NumContratoEsp", "IdentEsp", "NombresEsp", "ApellidosEsp",
                                "CargoEsp", "DependenciaEsp", "PerspectivaEsp", "ObjEstraEsp", "ObjAreaEsp", "MetaIndiEsp", "IndicadorEsp",
                                "TipoEsp", "UnidadEsp", "UnoEsp", "DosEsp", "TresEsp", "CuatroEsp", "PesoEsp", "Ident2Esp", "Nom2Esp", "Apell2Esp",
                                "Cargo2Esp", "Depen2Esp"
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
                                    ///




                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6]

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_AREA']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAño']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNRO_CONT']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAPE_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CARG']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_PERS']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_OBES']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_OBAR']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KtxtDesMeta_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_INDI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTIP_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblUNI_MEDI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_PTRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_STRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_TTRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_CTRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPOR_PESO']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAPE_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CARGE']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_ARBOE']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Gerencia u Oficina", "Año", "Numero de Contrato", "Identificación", "Nombres", "Apellidos",
                                        "Cargo", "Dependencia", "Perspectiva", "Objetivo Estratégico", "Objetivo Áreas", "Meta Individual",
                                        "Indicador", "Tipo", "Unidad", "Seguimiento I", "Seguimiento II", "Seguimiento III", "Seguimiento IV",
                                        "Peso %", "Identificación 2", "Nombres 2", "Apellidos 2", "Cargo 2", "Dependencia 2"
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
                                        if (counter == 11) for (int i = 0; i < 11; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 20)
                                        {
                                            for (int i = 0; i < 11; i++) SendKeys.Send("{UP}");
                                            selenium.Click("//a[contains(@id,'nav-evaluador-tab')]");
                                            Thread.Sleep(100);
                                            selenium.Screenshot("Segunda Pestaña", true, file);

                                        }

                                    }

                                    //No se puede usar por la repeticion de campo (Diccionario no admite llave duplicada)
                                    //Console.WriteLine(CamposPagina.Count);
                                    //elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    //foreach (var campo in elementos)
                                    //{

                                    //    if (campo.Key != campo.Value)
                                    //    {
                                    //        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                    //    }

                                    //}

                                    for (int i = 0; i < CamposPagina.Count; i++)
                                    {
                                        if (CamposMTM[i] != CamposPagina[i])
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + CamposMTM[i] + " y el encontrado es: " + CamposPagina[i]);
                                        }
                                    }


                                    /////////// Validación TABS ////////
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    //for (int i = 0; i < 16; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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

                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnAplicar')]");
                                    Thread.Sleep(500);
                                    //xpath=//a[@id='btnGuardar']
                                    selenium.Screenshot("Campos Necesarios", true, file);

                                    //Relative: xpath=//div[@id='printable']/div[17]/div[3]
                                    //Position: xpath=//div[17]/div[3]



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
        public void SmartPeople_frmEdFomeiFEmerNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmEdFomeiFEmerNTC")
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
                                "HomeEsp", "GereOfEsp", "AñoEsp", "NumContratoEsp", "IdentEsp", "NombresEsp", "ApellidosEsp",
                                "CargoEsp", "DependenciaEsp", "PerspectivaEsp", "ObjEstraEsp", "ObjAreaEsp", "MetaIndiEsp", "IndicadorEsp",
                                "TipoEsp", "UnidadEsp", "SeguimientoEsp",  "Ident2Esp",
                                "Nom2Esp", "Apell2Esp",  "Cargo2Esp", "Depen2Esp"
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

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6]

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_AREA']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAño']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNRO_CONT']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAPE_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CARG']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDependencia']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_PERS']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_OBES']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_OBAR']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDES_META']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_INDI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTIP_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblUNI_MEDI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTrimestre']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_PTRI']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_STRI']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_TTRI']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_CTRI']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblPOR_PESO']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAPE_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CARGE']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_ARBOE']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Gerencia u Oficina", "Año", "Numero de Contrato", "Identificación", "Nombres", "Apellidos",
                                        "Cargo", "Dependencia", "Perspectiva", "Objetivo Estratégico", "Objetivo Áreas", "Meta Individual",
                                        "Indicador", "Tipo", "Unidad", "Seguimiento", "Identificación 2", "Nombres 2", "Apellidos 2", "Cargo 2", "Dependencia 2"
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
                                        if (counter == 10) for (int i = 0; i < 12; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 16)
                                        {
                                            for (int i = 0; i < 12; i++) SendKeys.Send("{UP}");
                                            selenium.Click("//a[contains(text(),'Datos evaluador')]");
                                            Thread.Sleep(100);
                                            selenium.Screenshot("Segunda Pestaña", true, file);

                                        }

                                    }

                                    //No se puede usar por la repeticion de campo (Diccionario no admite llave duplicada)
                                    //Console.WriteLine(CamposPagina.Count);
                                    //elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    //foreach (var campo in elementos)
                                    //{

                                    //    if (campo.Key != campo.Value)
                                    //    {
                                    //        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                    //    }

                                    //}

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
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    //for (int i = 0; i < 16; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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

                                    for (int i = 0; i < 3; i++) SendKeys.Send("{DOWN}");
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnAplicar')]");
                                    Thread.Sleep(500);
                                    //xpath=//a[@id='btnGuardar']
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
        public void SmartPeople_frmEdFomeiRFNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmEdFomeiRFNTC")
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
                                "HomeEsp", "GereOfEsp", "AñoEsp", "NumContratoEsp", "IdentEsp", "NombresEsp", "ApellidosEsp",
                                "CargoEsp", "DependenciaEsp", "TipoDocEsp", "PerspectivaEsp", "ObjEstraEsp", "ObjAreaEsp", "MetaIndiEsp", "IndicadorEsp",
                                "TipoEsp", "UnidadEsp", "SeguimientoEsp", "UnoEsp", "DosEsp", "TresEsp", "CuatroEsp", "PesoEsp", "Ident2Esp",
                                "Nom2Esp", "Apell2Esp",  "Cargo2Esp", "Depen2Esp"
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

                                    selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                    Thread.Sleep(500);

                                    ////a[contains(text(),'Detalle')]
                                    ///


                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6]

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_AREA']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAño']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNRO_CONT']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAPE_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CARG']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrTipoDocumento_lblTIP_DOCU']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_PERS']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_OBES']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_OBAR']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDES_META']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_INDI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTIP_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblUNI_MEDI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTrimestre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_PTRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_STRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_TTRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_CTRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPOR_PESO']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAPE_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CARGE']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_ARBOE']",


                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Gerencia u Oficina", "Año", "Numero de Contrato", "Identificación", "Nombres", "Apellidos",
                                        "Cargo", "Dependencia", "Tipo de documento", "Perspectiva", "Objetivo Estratégico", "Objetivo Áreas", "Meta Individual",
                                        "Indicador", "Tipo", "Unidad", "Seguimiento", "I", "II", "III", "IV",
                                        "Peso %", "Identificación 2", "Nombres 2", "Apellidos 2", "Cargo 2", "Dependencia 2"
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
                                        if (counter == 8) for (int i = 0; i < 13; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 9) for (int i = 0; i < 11; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 22)
                                        {
                                            for (int i = 0; i < 24; i++) SendKeys.Send("{UP}");
                                            selenium.Click("//a[contains(@id,'nav-evaluador-tab')]");
                                            Thread.Sleep(100);
                                            selenium.Screenshot("Segunda Pestaña", true, file);

                                        }

                                    }

                                    //No se puede usar por la repeticion de campo (Diccionario no admite llave duplicada)
                                    //Console.WriteLine(CamposPagina.Count);
                                    //elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    //foreach (var campo in elementos)
                                    //{

                                    //    if (campo.Key != campo.Value)
                                    //    {
                                    //        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                    //    }

                                    //}

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
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    //for (int i = 0; i < 16; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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

                                    for (int i = 0; i < 4; i++) SendKeys.Send("{DOWN}");
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnAplicar')]");
                                    Thread.Sleep(500);
                                    //xpath=//a[@id='btnGuardar']
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
        public void SmartPeople_frmFdSoproNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmFdSoproNTC")
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
                                "HomeEsp", "GuardarEsp", "ConsecEsp", "NomProyEsp", "NivelCapaEsp", "AreaCoEsp", "NumParEsp",
                                "IntenHorEsp", "PerDuEsp", "FechaInEsp", "FechaFinEsp", "CiudadEsp", "LugarEsp", "RespoEsp", "OtraEntEsp",
                                "DescEsp", "CargosEsp", "DepenEsp", "JustiEsp", "ObjGenEsp", "ObjEspeEsp", "AlinEsp", "PlantEsp", "PerfDocEsp",
                                "PerfFormEsp", "MetodEsp", "LogiEsp"
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
                                    ///




                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Guardar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        [Guardar] = CamposTotalesMTM[7]

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label1']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label3']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label11']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNumParc']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblIntHora']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecini_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecfin_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrDivi_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblLugProy']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label8']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblOtrEnti']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_txtDesNece_lblTexto']", //Descripcion
                                        "xpath=//span[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel2_TabContainer2_TabPane11_lblCodCarg']", //cargo
                                        "xpath=//span[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel2_TabContainer2_TabPane22_lblDependencias']", //Dependencias
                                        "xpath=//span[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel3_txtJusTifi_lblTexto']",//Justificacion
                                        "xpath=//span[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel4_txtObjGene_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel4_txtObjEspe_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel5_txtAliProy_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel6_txtPlaAlte_lblTexto']", //planteamiento
                                        "xpath=//span[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel7_txtPerDisc_lblTexto']", //Perfil docente
                                        "xpath=//span[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel7_txtPerForm_lblTexto']", //Perfil de formador
                                        "xpath=//span[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel10_txtMetOdol_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel11_txtLogIsti_lblTexto']",


                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Consecutivo", "Nombre del Proyecto", "Nivel de Capacitación", "Area de Conocimiento", "Número de Participante",
                                        "Intensidad Horaria", "Periodo de Duración", "Fecha Inicial", "Fecha Final", "Ciudad", "Lugar", "Responsable", "Otra Entidad",
                                        "Descripción", "Cargos", "Dependencias", "Justificación", "Objetivos Generales", "Objetivos Especificos", "Alineación",
                                        "Planteamiento", "Perfiles del Docente", "Perfiles del Formador", "Metodología", "Logistica"
                                    };

                                    for (int i = 0; i < xpath.Count; i++)
                                    {
                                        xpath[i] = xpath[i].Replace("xpath=", "");
                                        xpath[i] = xpath[i].Replace("=", ",");
                                        xpath[i] = xpath[i].Insert(7, "contains(");
                                        xpath[i] = xpath[i].Insert(xpath[i].Length - 1, ")");
                                        campos.Add(xpath[i], descripcionSC[i]);
                                    }

                                    //Debugger.Launch();
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
                                        if (counter == 9) for (int i = 0; i < 12; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 14) selenium.Click("//a[contains(text(),'Cubrimiento')]");
                                        if (counter == 15) selenium.Click("//a[contains(text(),'Dependencias')]");
                                        if (counter == 16) selenium.Click("//a[contains(text(),'Justificación')]");
                                        if (counter == 17)
                                        {
                                            selenium.Click("//a[contains(text(),'Objetivos')]");
                                            for (int i = 0; i < 4; i++) SendKeys.Send("{UP}");
                                        }
                                        if (counter == 19)
                                        {
                                            for (int i = 0; i < 2; i++) SendKeys.Send("{UP}");
                                            selenium.Click("//a[contains(text(),'Alineación')]");
                                        }

                                        if (counter == 20) selenium.Click("//a[contains(text(),'Plantemiento')]");
                                        if (counter == 21)
                                        {

                                            selenium.Click("//a[contains(text(),'Perfiles')]");

                                            selenium.Click("//a[contains(@id,'__tab_ctl00_ContenidoPagina_TabContainer1_TabPanel7')]/span");
                                            for (int i = 0; i < 4; i++) SendKeys.Send("{UP}");
                                        }
                                        if (counter == 23) selenium.Click("//a[contains(@id,'__tab_ctl00_ContenidoPagina_TabContainer1_TabPanel10')]/span");
                                        if (counter == 24) selenium.Click("//a[contains(@id,'__tab_ctl00_ContenidoPagina_TabContainer1_TabPanel11')]/span");
                                    }

                                    //No se puede usar por la repeticion de campo (Diccionario no admite llave duplicada)
                                    //Console.WriteLine(CamposPagina.Count);
                                    //elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    //foreach (var campo in elementos)
                                    //{

                                    //    if (campo.Key != campo.Value)
                                    //    {
                                    //        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                    //    }

                                    //}

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
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    for (int i = 0; i < 12; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
        public void SmartPeople_frmLIEdFomeiANTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmLIEdFomeiANTC")
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
                                "HomeEsp", "GereOfEsp", "AñoEsp", "NumContratoEsp", "IdentEsp", "NombresEsp", "ApellidosEsp",
                                "CargoEsp", "DependenciaEsp", "PerspectivaEsp", "ObjEstraEsp", "ObjAreaEsp", "MetaIndiEsp", "IndicadorEsp",
                                "TipoEsp", "UnidadEsp", "SeguimientoEsp", "UnoEsp", "DosEsp", "TresEsp", "CuatroEsp", "PesoEsp", "Ident2Esp",
                                "Nom2Esp", "Apell2Esp",  "Cargo2Esp", "Depen2Esp"
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



                                    selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                    Thread.Sleep(500);


                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    //string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        //[Guardar] = CamposTotalesMTM[7]

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_AREA']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAño']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNRO_CONT']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAPE_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CARG']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_PERS']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_OBES']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_OBAR']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDES_META']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_INDI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTIP_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblUNI_MEDI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTrimestre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_PTRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_STRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_TTRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_CTRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPOR_PESO']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAPE_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CARGE']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_ARBOE']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Gerencia u Oficina", "Año", "Numero de Contrato", "Identificación", "Nombres", "Apellidos",
                                        "Cargo", "Dependencia", "Perspectiva", "Objetivo Estratégico", "Objetivo Áreas", "Meta Individual",
                                        "Indicador", "Tipo", "Unidad", "Seguimiento", "I", "II", "III", "IV",
                                        "Peso %", "Identificación 2", "Nombres 2", "Apellidos 2", "Cargo 2", "Dependencia 2"
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
                                        if (counter == 8) for (int i = 0; i < 17; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 13) for (int i = 0; i < 4; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 21)
                                        {
                                            for (int i = 0; i < 18; i++) SendKeys.Send("{UP}");
                                            selenium.Click("//a[contains(@id,'nav-evaluador-tab')]");
                                        }
                                        if (counter == 26) for (int i = 0; i < 15; i++) SendKeys.Send("{UP}");
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
                                    //selenium.Click("//div[@id='ctl00_pBotones']/div");

                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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

                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnAplicar')]");
                                    Thread.Sleep(500);
                                    //xpath=//a[@id='btnGuardar']
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
        public void SmartPeople_frmLIEdFomeiTNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmLIEdFomeiTNTC")
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
                                "HomeEsp", "GereOfEsp", "AñoEsp", "NumContratoEsp", "IdentEsp", "NombresEsp", "ApellidosEsp",
                                "CargoEsp", "DependenciaEsp", "PerspectivaEsp", "ObjEstraEsp", "ObjAreaEsp", "MetaIndiEsp", "IndicadorEsp",
                                "TipoEsp", "UnidadEsp", "SeguimientoEsp", "UnoEsp", "DosEsp", "TresEsp", "CuatroEsp", "PesoEsp", "Ident2Esp",
                                "Nom2Esp", "Apell2Esp",  "Cargo2Esp", "Depen2Esp"
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

                                    selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                    Thread.Sleep(500);




                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    //string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        //[Guardar] = CamposTotalesMTM[7]

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_AREA']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAño']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNRO_CONT']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAPE_EMPL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CARG']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_PERS']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_OBES']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_OBAR']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDES_META']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_INDI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTIP_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblUNI_MEDI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTrimestre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_PTRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_STRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_TTRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPRO_CTRI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPOR_PESO']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNOM_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAPE_EVAL']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_CARGE']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCOD_ARBOE']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Gerencia u Oficina", "Año", "Numero de Contrato", "Identificación", "Nombres", "Apellidos",
                                        "Cargo", "Dependencia", "Perspectiva", "Objetivo Estratégico", "Objetivo Áreas", "Meta Individual",
                                        "Indicador", "Tipo", "Unidad", "Seguimiento", "I", "II", "III", "IV",
                                        "Peso %", "Identificación 2", "Nombres 2", "Apellidos 2", "Cargo 2", "Dependencia 2"
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
                                        if (counter == 8) for (int i = 0; i < 18; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 21)
                                        {
                                            for (int i = 0; i < 18; i++) SendKeys.Send("{UP}");
                                            selenium.Click("//a[contains(@id,'nav-evaluador-tab')]");
                                            Thread.Sleep(100);
                                            selenium.Screenshot("Segunda Pestaña", true, file);

                                        }
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
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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




                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnAplicar')]");
                                    Thread.Sleep(500);
                                    //xpath=//a[@id='btnGuardar']
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
        public void SmartPeople_frmLINmAplvaNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmLINmAplvaNTC")
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
                                "HomeEsp", "GuardarEsp", "NumContEsp", "IdentEsp",
                                "NomApeEsp", "CodInterEsp", "CargoEsp", "AplazEsp", "ConcAsoEsp", "SecuencialEsp", "PeriodoEsp", "FechaIniCaEsp",
                                "FechaFinCaEsp", "FechaIniDisEsp", "FechaFinDisEsp", "DiasVacEsp", "Aplaz2Esp", "TipoNovEsp", "TipoDiasEsp",
                                "ReanudeEsp", "FechaAplazInEsp", "DiasAplazEsp", "DiasTomaEsp", "PendiEsp", "FechaIniEsp", "FechaFinEsp", "ObservEsp", "EstadoEsp"
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
                                    ///




                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Guardar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Guardar = selenium.EmergenteBotones("btnGuardar");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        [Guardar] = CamposTotalesMTM[7]

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomCarg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblApla']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodConc']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSecEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPerDisf']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecCau1']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecCau2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecInid']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecFind']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDiaTomt']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSecApla']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipDias']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblIndRean']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecApla_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDiaApla']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDiaToma']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDiaPend']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecDisd_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecDish_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrltxtObsErva_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label8']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Nro. Contrato", "Identificación",
                                        "Nombres y Apellidos", "Código Interno", "Cargo", "Aplazamiento", "Concepto asociado", "Secuencial",
                                        "Periodo", "Fecha inicial causación", "Fecha final causación", "Fecha inicial disfrute", "Fecha final disfrute",
                                        "Días vacaciones", "Aplazamiento", "Tipo de Novedad", "Tipo de días", "Reanude", "Fecha de apla", "Dias apla",
                                        "Días toma", "Pendientes", "Fecha Inicial", "Fecha Final", "Observaciones", "Estado"
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
                                        if (counter == 19) for (int i = 0; i < 13; i++) SendKeys.Send("{DOWN}");
                                    }

                                    //No se puede usar por la repeticion de campo (Diccionario no admite llave duplicada)
                                    //Console.WriteLine(CamposPagina.Count);
                                    //elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    //foreach (var campo in elementos)
                                    //{

                                    //    if (campo.Key != campo.Value)
                                    //    {
                                    //        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                    //    }

                                    //}

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
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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

                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(500);
                                    //xpath=//a[@id='btnGuardar']
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
        public void SmartPeople_frmLISlReqpeNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmLISlReqpeNTC")
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
                                "HomeEsp", "RetornarEsp", "GuardarEsp", "GrupoReqEsp", "FormCobEsp", "FiltSeleEsp",
                                "ValiVacEsp", "FechaSolEsp", "CargProEsp", "NumPlaEsp", "FechaPoInEsp", "CenCostEsp", "MotivSolEsp",
                                "TipoConEsp", "SueldoBaEsp", "CiudadEsp", "DetReqEsp", "ObsSolEsp", "PubWebEsp", "DocuEsp"
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
                                    ///




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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodGrse']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblForCobe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFilSele']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValVaca']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaSoli_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroPlaz']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaPoin_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCcos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodMoti']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSueProp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrDivipEMPR_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KtxtDetRequ_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KtxtObserSoli_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblVisSuew']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblArcAdju']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Grupo de Requisiciones", "Forma de Cobertura de la Vacante", "Filtro de Selección", "Valida Vacante",
                                        "Fecha de Solicitud", "Cargo a Proveer", "N. de Plazas", "Fecha Posible Ingreso", "Centro de Costo", "Motivo de la Solicitud",
                                        "Tipo de Contrato", "Sueldo Básico", "Ciudad", "Detalle de la Requisición", "Observaciones de la Solicitud", "Publicar Sueldo Web",
                                        "Documentos"
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
                                        if (counter == 9) for (int i = 0; i < 2; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 12) for (int i = 0; i < 19; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 16) for (int i = 0; i < 10; i++) SendKeys.Send("{DOWN}");
                                    }

                                    //No se puede usar por la repeticion de campo (Diccionario no admite llave duplicada)
                                    //Console.WriteLine(CamposPagina.Count);
                                    //elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    //foreach (var campo in elementos)
                                    //{

                                    //    if (campo.Key != campo.Value)
                                    //    {
                                    //        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                    //    }

                                    //}

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
                                    for (int i = 0; i < 30; i++) SendKeys.Send("{UP}");
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
        public void SmartPeople_frmLISlReqpeDianNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmLISlReqpeDianNTC")
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
                                "HomeEsp", "RetornarEsp", "GuardarEsp", "GrupoReqEsp", "FormCobEsp", "FiltSeleEsp", "ValiVacEsp",
                                "FechaSolEsp", "TipVinEsp", "ProceEsp", "NivelEsp", "CargProEsp", "FechaPoInEsp", "CenCostEsp", "NumPlazaEsp",
                                "MotivSolEsp", "TipConEsp", "SuelBaEsp", "UbiEsp", "CiudadEsp", "ObserEsp", "PubWebEsp"
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
                                    ///




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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodGrse']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblForCobe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFilSele']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValVaca']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaSoli_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lbltipVinc']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblProce']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNivel']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaPoin_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCcos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroPlaz']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodMoti']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSueProp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblUbicacion']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlDivPoli_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KtxtObserSoli_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblVisSuew']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Grupo de Requisiciones", "Forma de Cobertura de la Vacante", "Filtro de Selección", "Validar Vacantes",
                                        "Fecha de Solicitud", "Tipo de Vinculación", "Proceso", "Nivel", "Cargo a Proveer", "Fecha Posible Ingreso",
                                        "Centro Costo", "N. de Plazas", "Motivo de la Solicitud", "Tipo de Contrato", "Sueldo Básico", "uUbicación",
                                        "Ciudad", "Observaciones de la Solicitud", "Publicar Sueldo Web"
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
                                        if (counter == 12) for (int i = 0; i < 15; i++) SendKeys.Send("{DOWN}");
                                    }

                                    //No se puede usar por la repeticion de campo (Diccionario no admite llave duplicada)
                                    //Console.WriteLine(CamposPagina.Count);
                                    //elementos = CamposPagina.Zip(CamposMTM, (k, v) => new { Key = k, Value = v }).ToDictionary(x => x.Key, x => x.Value);

                                    //foreach (var campo in elementos)
                                    //{

                                    //    if (campo.Key != campo.Value)
                                    //    {
                                    //        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + campo.Value + " y el encontrado es: " + campo.Key);
                                    //    }

                                    //}

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
                                    for (int i = 0; i < 15; i++) SendKeys.Send("{UP}");
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
        public void SmartPeople_frmLISlReqpeHNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmLISlReqpeHNTC")
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
                                "HomeEsp", "RetornarEsp", "GuardarEsp", "CasoEsp", "AreaUsuEsp", "FormCobEsp", "FiltSeleEsp",
                                "ValiVacEsp", "FechaSolEsp", "CargProEsp", "NumPlazasEsp", "FechaPoInEsp", "CentCostEsp", "CentTrabEsp",
                                "HoraTrabEsp", "JefeInmeEsp", "DereDotEsp", "PagoComEsp", "PorcenConEsp", "PagoGaranEsp", "PorcenGaranEsp",
                                "TiempGaranEsp", "OtrosPagEsp", "ValOtrosPagEsp", "TipSalEsp", "SuelBaEsp", "HoraExFiEsp",
                                "MotivSolEsp", "TipConEsp", "FunRemEsp", "CiudadEsp", "ObserSoliEsp", "ReqCargEsp", "ActFiEsp", "RopaCalEsp",
                                "IdentEsp", "SisInEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCasCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodGrse']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblForCobe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFilSele']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValVaca']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaSoli_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroPlaz']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaPoin_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCcos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCenp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodTurn']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodFrep']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDerDota']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPagComi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label12']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPagGara']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValGara']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTieGara']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblOtrPago']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValOtro']",
                                        //"xpath=//span[@id='ctl00_ContenidoPagina_lblNomOtpa']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipSala']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSueProp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblHraExtf']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodMoti']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodReem']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlDivPoli_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KtxtObserSoli_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRequiRecu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRequA']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblRopaN']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblIdentI']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblIdentO']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Caso", "Área Usuaria", "Forma de Cobertura de la Vacante", "Filtro de Selección",
                                        "Validar Vacantes en", "Fecha de Solicitud", "Cargo a Proveer", "N. de Plazas", "Fecha Posible Ingreso",
                                        "Centro de Costo", "Centro de Trabajo", "Horario de Trabajo", "Jefe Inmediato", "Derecho a Dotación",
                                        "Pago Comisiones", "Porcentaje Comisión", "Pago Garantizado", "Procentaje Garantizado", "Tiempo Garantizado",
                                        "Otros Pagos", "Valor Otros Pagos",  "Tipo de Salario", "Sueldo Básico", "Horas extras fijas",
                                        "Motivo de la Solicitud", "Tipo de Contrato", "Funcionario a reemplazar", "Ciudad", "Observaciones de la Solicitud",
                                        "Requisitos de recursos del cargo", "ACTIVOS FIJOS", "ROPA Y CALZADO DE LABOR", "IDENTIFICACIÓN", "SISTEMAS DE INFORMACIÓN"
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
                                        if (counter == 10) for (int i = 0; i < 13; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 28) for (int i = 0; i < 14; i++) SendKeys.Send("{DOWN}");

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
                                    for (int i = 0; i < 23; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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

                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(1500);
                                    SendKeys.Send("{ENTER}");
                                    Thread.Sleep(1500);
                                    SendKeys.Send("{ENTER}");

                                    // selenium.Screenshot("Campos Necesarios", true, file);






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
        public void SmartPeople_frmLISlReqpeTemNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmLISlReqpeTemNTC")
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
                                "HomeEsp", "RetornarEsp", "GuardarEsp", "GrupRecEsp", "FormCobEsp", "FiltSeleEsp",
                                "FechaSolEsp", "CentCosEsp", "NumPlaza", "CargProEsp", "FechaPoInEsp", "MotivSolEsp", "PubWebEsp",
                                "ContratoEsp", "SubTipConEsp", "DetReqEsp", "ObSolEsp", "CiudadEsp", "FunRemEsp", "DocuEsp", "TipDocEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodGrse']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblForCobe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFilSele']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaSoli_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCcos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroPlaz']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaPoin_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodMoti']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblVisSuew']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipContr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KtxtDetRequ_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KtxtObserSoli_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlDivPoli_lblDivPoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodReem']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblArcAdju']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrTipoDocumento_lblTIP_DOCU']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Grupo de requisiciones", "Forma de cobertura de la vacante", "Filtro de selección",
                                        "Fecha de solicitud", "Centro de costo", "Número de plazas", "Cargo a proveer", "Fecha posible ingreso",
                                        "Motivo de la solicitud", "Publicar sueldo web", "Contrato", "Sub-tipo de contrato", "Detalle de la requisición",
                                        "Observaciones de la solicitud", "Ciudad", "Funcionario a reemplazar", "Documentos", "Tipo de documento"
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
                                        if (counter == 9) for (int i = 0; i < 14; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 15) for (int i = 0; i < 10; i++) SendKeys.Send("{DOWN}");
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
                                    for (int i = 0; i < 25; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
        public void SmartPeople_frmNmCtperNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmNmCtperNTC")
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
                                "HomeEsp", "RetornarEsp", "GuardarEsp","FechaSolEsp", "MotivAuEsp", "EstadoEsp", "FechaIniPerEsp",
                                "HoraIniEsp", "FechaFinPerEsp", "HoraFinEsp", "ObsPerEsp", "ObsSolEsp", "DocSolEsp", "TipoDocEsp", "IdentEsp", "NomApelEsp", "NumCont", "DepenEsp",
                                "ArbolEsp", "CargoEsp", "CenCostEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_txtFecSoli_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodMaus']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstPerm']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_kcfFechaIniPermiso_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlHorSali_lblHora']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_kcfFechaFinPermiso_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlHorEntr_lblHora']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_txtObsErva_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_txtObsErvaSoli_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblArcAdju']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrTipoDocumento_lblTIP_DOCU']",

                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNmbEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDependencia']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblArbol']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCcos']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Fecha de Solicitud", "Motivo del Ausentismo", "Estado", "Fecha Inicio de Permiso",
                                        "Hora Inicio", "Fecha Fin de Permiso", "Hora Fin", "Observaciones de Permisos", "Observaciones de la Solicitud",
                                        "Documento de Soporte", "Tipo de documento","Identificación", "Nombres y Apellidos", "Nro. Contrato", "Dependencia", "Árbol", "Cargo",
                                        "Centro de Costo"
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

                                        if (counter == 9) for (int i = 0; i < 5; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 11)
                                        {
                                            for (int i = 0; i < 6; i++) SendKeys.Send("{UP}");
                                            selenium.Click("//div[@id='headingOne']/h2/button/span");

                                        }

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
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
        public void SmartPeople_frmNmCtperDNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmNmCtperDNTC")
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
                                "HomeEsp", "FechaSolEsp", "MotivAuEsp", "EstadoEsp", "FechaIniPerEsp",
                                "HoraIniEsp", "FechaFinPerEsp", "HoraFinEsp", "ObsPerEsp", "ObsSolEsp", "IdentEsp", "NomApelEsp", "NumCont", "DepenEsp",
                                "ArbolEsp", "CargoEsp", "CenCostEsp"
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


                                    //titulo de la pagina
                                    Screenshot(CamposTotalesMTM[4], true, file);

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6]

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


                                    //ERROR
                                    try
                                    {
                                        selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                        Thread.Sleep(500);
                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }



                                    List<string> xpath = new List<string>() {

                                         "xpath=//span[@id='ctl00_ContenidoPagina_txtFecSoli_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodMaus']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEstPerm']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_kcfFechaIniPermiso_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlHorSali_lblHora']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_kcfFechaFinPermiso_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlHorEntr_lblHora']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_txtObsErva_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_txtObsErvaSoli_lblTexto']",

                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNmbEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDependencia']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblArbol']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCcos']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Fecha de Solicitud", "Motivo del Ausentismo", "Estado", "Fecha Inicio de Permiso",
                                        "Hora Inicio", "Fecha Fin de Permiso", "Hora Fin", "Observaciones de Permisos", "Observaciones de la Solicitud","Identificación", "Nombres y Apellidos", "Nro. Contrato", "Dependencia", "Árbol", "Cargo",
                                        "Centro de Costo"

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
                                        ////*[@id="headingOne"]/h2[1]/button[1]

                                        if (counter == 9)
                                        {

                                            selenium.Click("//div[@id='headingOne']/h2/button");

                                        }

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

                                    // Debugger.Launch();

                                    /////////// Validación TABS ////////
                                    //selenium.Click(xpath[0]);
                                    //  for (int i = 0; i < 11; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0) ////*[@id="ctl00_ContenidoPagina_txtFecSoli_txtFecha"]
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
                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']
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
        public void SmartPeople_frmNmRcappANTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmNmRcappANTC")
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
                                "HomeEsp", "NumConEsp", "IdentEsp", "NomApellEsp", "CodInterEsp", "FechaCorteEsp", "FechaSolEsp",
                                "NumMesEsp", "LinCreEsp", "TasaInEsp", "PlazoEsp", "MontoEsp", "InteMoEsp",
                                "InterQuinEsp","AnualEsp", "QuinEsp", "TotalQuinEsp"
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

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6]

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


                                    //ERROR
                                    try
                                    {
                                        selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                        Thread.Sleep(500);
                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }

                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecCort_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecSoli_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMeses']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblLinCred']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTasInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPlazo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMonto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblInteresesMonto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblInteresQuincena']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblAnual']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblQuincenal']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTotQuince']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Nro Contrato", "Identificación", "Nombres y Apellidos", "Cod. Interno", "Fecha Corte", "Fecha Solicitud",
                                        "Numero de meses", "Lineas de Credito", "Tasa Interes", "Plazo", "Monto", "Interes Monto", "Interes Quicena", "Anual",
                                        "Quincenal", "Total Quincena"
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
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']
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
        public void SmartPeople_frmNmSanciDNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmNmSanciDNTC")
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
                                "HomeEsp", "IdentEsp", "NomApelEsp", "NumConEsp", "CargoEsp", "CenCostEsp", "FaltaEsp",
                                "FechaHeEsp", "FechaDesEsp", "HechosEsp", "EvalJuEsp", "DesJuEsp", "FechaSanEsp", "NumSanEsp",
                                "FechaCitEsp", "DesEmpEsp", "FechaDes2Esp", "FechaNotEsp", "ClaseNomEsp", "NomServEsp",
                                "FunRepEsp", "RespSegEsp", "FechaAsigEsp", "TiRetaEsp", "TiAusenEsp", "ItemCritEsp", "ConNovEsp",
                                "FechaDesdeEsp", "FechaHasta", "DiasEsp", "GeneNovEsp", "AfectarNoEsp", "AfectarNoveEsp", "ObsEsp"
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

                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6]

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


                                    //ERROR
                                    try
                                    {
                                        selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                        Thread.Sleep(500);
                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }

                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCcos']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodFalt']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecHech_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecDesc_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrltxtTexSanc_lblTexto']", //hechos
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrltxtEvaJuri_lblTexto']", //Evaluacion juridica
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrltxtDesJuri_lblTexto']",//desicion juridica
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecReso_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNumReso']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecCita_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodDeem']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecDeci_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecNoti_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodTnom']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomServ']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblGteDpto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblResPons']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecAsig_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTieReta']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTieAuse']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblIteCrit']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodConc']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecDesd_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecHast_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNumDias']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblGenNove']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecAfno_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecNove_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblObsErva']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Identificación", "Nombres y Apellidos", "Nro. Contrato", "Cargo",
                                        "Centro de Costo", "Falta", "Fecha Hechos", "Fecha Descargo", "Hechos",
                                        "Evaluación Juridica", "Decisión Juridica", "Fecha Sanción", "Nro Sanción",
                                        "Fecha Citación", "Decisión Empresa", "Fecha Decisión", "Fecha Notificación",
                                        "Clase de Nómina", "Nombre Servicio", "Funcionario quien Reporta", "Responsable del Seguimiento",
                                        "Fecha Asignación", "Tiempo Retardo(minutos)", "Tiempo Ausencia(días)", "Ítem Critico",
                                        "Concepto Novedades", "Fecha Desde", "Fecha Hasta", "Días", "Genera Novedad", "Afectar Nómina a partir de",
                                        "Afectar Novedad", "Observaciones"
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
                                        if (counter == 9) selenium.Click("//a[contains(text(),'Descargos')]");
                                        if (counter == 10) selenium.Click("//a[contains(text(),'Descisión Juridica')]");
                                        if (counter == 11) for (int i = 0; i < 15; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 17)
                                        {
                                            for (int i = 0; i < 3; i++) SendKeys.Send("{DOWN}");
                                            selenium.Click("//a[contains(text(),'Procesos Disciplinarios')]");
                                        }
                                        if (counter == 25)
                                        {
                                            selenium.Click("//a[contains(text(),'Nómina')]");
                                            Thread.Sleep(500);
                                            for (int i = 0; i < 3; i++) SendKeys.Send("{DOWN}");
                                        }

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
        public void SmartPeople_frmBpBeotoDNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmBpBeotoDNTC")
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
                                "HomeEsp", "RetornarEsp", "IdentEsp", "NomApelEsp", "NumContEsp", "CargoEsp", "BeneEsp",
                                "FamEsp", "FechaCorteEsp", "ValSolEsp", "ObsEsp", "DocResEsp", "ObsSolEsp", "DocuEsp"
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
                                        selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Error", true, file);

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

                                    var InitialCheck = new Dictionary<string, string>()
                                    {
                                        [Titulo] = CamposTotalesMTM[4],
                                        [Subtitulo] = CamposTotalesMTM[5],
                                        [Home] = CamposTotalesMTM[6],
                                        [Retornar] = CamposTotalesMTM[7]
                                    };


                                    int controlCounter = 0;
                                    foreach (var nameControl in InitialCheck)
                                    {
                                        Thread.Sleep(100);
                                        if (nameControl.Key.ToString() != nameControl.Value.ToString())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: Nombre de " + name[controlCounter] + " incorrecto, el esperado es: " + nameControl.Value + " y el encontrado es: " + nameControl.Key);
                                        }

                                        controlCounter = controlCounter + 1;
                                    }

                                    var campos = new Dictionary<string, string>();




                                    List<string> xpath = new List<string>() {
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodCarg']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodMces']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomFami']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFec_Regi_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValCesp']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlObserva_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDocReque']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KtxtObserSoli_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblArcAdju']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Identificación", "Nombres y Apellidos", "Número de Contrato", "Cargo", "Beneficios", "Familiar",
                                        "Fecha de Corte", "Valor Solicitado", "Observaciones", "Documentación Requerida",
                                        "Observaciones de la Solicitud", "Documentos"
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
                                        if (counter == 9) for (int i = 0; i < 15; i++) SendKeys.Send("{DOWN}");
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
                                    for (int i = 0; i < 15; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']
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
        public void SmartPeople_frmBpSolceNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmBpSolceNTC")
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
                                "HomeEsp", "GuardarEsp", "IdentEsp", "NomApelEsp", "NumContEsp", "FechaSolEsp",
                                "TipoSolEsp", "MotivSolEsp", "ObsEsp"
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
                                        selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Error", true, file);

                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }


                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Guardar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Guardar = selenium.EmergenteBotones("btnGuardar");
                                    List<string> nameBotEn = new List<string>() { Titulo, Subtitulo, Home, Guardar };

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecSoli_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipSoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtMotSoli_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtObsErva_lblTexto']",
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Identificación", "Nombres y Apellidos", "Número de Contrato", "Fecha de Solicitud",
                                        "Tipo de Solicitud", "Motivo de la solicitud", "Observaciones"
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
                                    //for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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

                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(500);
                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']
                                    try
                                    {
                                        selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Error", true, file);

                                    }
                                    catch (Exception e)
                                    {
                                        selenium.Screenshot("Campos Necesarios", true, file);

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
        public void SmartPeople_frmCoEncueNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmCoEncueNTC")
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
                                "HomeEsp", "RetornarEsp", "GuardarEsp", "FechaEnEsp", "ModuloEsp", "NumContEsp", "IdentEsp",
                                "NomEsp", "ApellEsp", "CodEsp", "EncueEsp", "VariEsp", "Nom2Esp", "PesoEsp", "DefDetEsp"
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
                                        selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Error", true, file);

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecEncu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblModulo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_LblId']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_LblNombre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblApe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodigo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEncuesta']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblVariable']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNom']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblPeso']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_txtDefVari_lblTexto']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Fecha Encuesta", "Modulo", "Numero Contrato", "Identificación", "Nombre",
                                        "Apellido", "Código", "Encuesta", "Variable", "Nombre", "Peso", "Definición Detallada"
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
                                    //for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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

                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(500);
                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']
                                    try
                                    {

                                        selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Error", true, file);



                                    }
                                    catch (Exception e)
                                    {
                                        selenium.Screenshot("Campos Necesarios", true, file);

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
        public void SmartPeople_frmCoEncueLVNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmCoEncueLVNTC")
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
                                "HomeEsp", "RetornarEsp", "NumContEsp", "IdentEsp", "NomEsp", "ApellEsp", "CodEsp",
                                "EncuEsp", "ModValEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_LblId']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_LblNombre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblApe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodigo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEncuesta']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divVari']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Numero de Contrato", "Identificación", "Nombre", "Apellido",
                                        "Código", "Encuesta", "Modulo Variable"
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
                                    //for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
        public void SmartPeople_frmCoEncueLVRNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmCoEncueLVRNTC")
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
                                "HomeEsp", "NumContEsp", "IdentEsp", "NomEsp", "ApellEsp", "CodEsp", "EncuEsp", "FechaIniEsp",
                                "FechaFinEsp", "PromEsp", "InterRanEsp", "PromVaEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_LblId']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_LblNombre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblApe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodigo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEncuesta']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFEC_INIC']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFEC_FINA']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblProTo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblInterRango']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divVari']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Número de Contrato", "Identificación", "Nombre", "Apellido", "Código",
                                        "Encuesta", "Fecha Inicio", "Fecha Final", "Promedio Total por Encuesta",
                                        "Interpretación del Rango", "Promedio de Variables"
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
                                    //for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
        public void SmartPeople_frmCoEncueTNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmCoEncueTNTC")
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
                                "HomeEsp", "GuardarEsp", "FechaEnEsp", "ModEsp", "NumContrato", "IdentEsp", "NomEsp",
                                "ApellEsp", "CodEsp", "EncuEsp", "PregEsp"
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


                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Guardar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Guardar = selenium.EmergenteBotones("btnGuardar");
                                    List<string> nameBotEn = new List<string>() { Titulo, Subtitulo, Home, Guardar };

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecEncu']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblModulo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_LblId']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_LblNombre']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblApe']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodigo']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblEncuesta']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_divPreguntas']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Fecha Encuesta", "Modulo", "Número Contrato", "Identificación",
                                        "Nombre", "Apellido", "Código", "Encuesta", "Preguntas"
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
                                    //for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']

                                    if (selenium.ExistControl("//a[contains(@id,'ctl00_btnCerrar')]"))
                                    {
                                        selenium.Screenshot("Error", true, file);

                                        Thread.Sleep(500);
                                        selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");

                                    }
                                    else
                                    {
                                        selenium.Screenshot("Campos Necesarios", true, file);

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
        public void SmartPeople_frmFdClsesNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmFdClsesNTC")
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
                                "HomeEsp", "RetornarEsp", "PlanEsp", "ProgramEsp", "CursoEsp", "FechaIniEsp", "FechaFinEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label1']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label3']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label8']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Plan", "Programa", "Curso", "Fecha Inicio", "Fecha Finalización"
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
                                    //for (int i = 0; i < 20; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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


                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnGuardar')]");
                                    Thread.Sleep(500);
                                    //SendKeys.Send("{ENTER}");
                                    //Thread.Sleep(500);
                                    //xpath =//a[@id='btnGuardar']

                                    if (selenium.ExistControl("//a[contains(@id,'ctl00_btnCerrar')]"))
                                    {
                                        selenium.Screenshot("Error", true, file);

                                        Thread.Sleep(500);
                                        selenium.Click("//a[contains(@id,'ctl00_btnCerrar')]");

                                    }
                                    else
                                    {
                                        selenium.Screenshot("Campos Necesarios", true, file);

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
        public void SmartPeople_frmNmAntcaNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmNmAntcaNTC")
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
                                "HomeEsp", "GuardarEsp", "TipOperEsp", "IdentEsp", "NomApelEsp", "NumContEsp", "FechaViEsp",
                                "HoraViEsp", "FechaRegreEsp", "HoraRegreEsp", "ClientEsp", "DetClientEsp", "SucurEsp",
                                "ProyEsp", "AreaEsp", "CentCostEsp", "ValSolEsp", "DesEsp", "ObsEsp"
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


                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Guardar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Guardar = selenium.EmergenteBotones("btnGuardar");
                                    List<string> nameBotEn = new List<string>() { Titulo, Subtitulo, Home, Guardar };

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label2']", //Nom Apel
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDesSoli']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblObsSoli']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Tipo de Operación", "Identificación", "Nombres y Apellidos", "Numero Contrato",
                                        "Fecha de Viaje", "Hora de Viaje", "Fecha de Regreso", "Hora de Regreso",
                                        "Cliente", "Detalle Cliente", "Sucursal", "Proyecto", "Area", "Centro de Costos",
                                        "Valor Solicitados", "Descripción", "Observaciones"
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
                                        if (counter == 15) for (int i = 0; i < 8; i++) SendKeys.Send("{DOWN}");
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
                                    for (int i = 0; i < 8; i++) SendKeys.Send("{UP}");
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    //List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
        public void SmartPeople_frmNmApldeNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmNmApldeNTC")
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
                                "HomeEsp", "NumContEsp", "IdentEsp", "NomApelEsp", "CodInEsp", "FechaIniEsp", "FechaFinEsp",
                                "DiasAplaEsp", "DiasTomEsp", "PendEsp", "FechaResEsp", "NumResEsp", "ObsEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecDisd_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecDish_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDiaApla']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDiaToma']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDiaPend']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecReso_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroReso']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtValMObsErva_lblTexto']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Numero de Contrato", "Identificación", "Nombres y Apellidos", "Cod. Interno",
                                        "Fecha Inicial", "Fecha Final", "Días Aplazados", "Días Tomados", "Pendientes",
                                        "Fecha de Resolución", "Nro. Resolución", "Observaciones"

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
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
        public void SmartPeople_frmNmAplvaDNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmNmAplvaDNTC")
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
                                "HomeEsp", "NumContEsp", "IdentEsp", "NomApelEsp", "CodInterEsp", "AplazEsp", "ConAsoEsp",
                                "SecuenEsp", "Aplaz2Esp", "FechaIniEsp", "FechaFinEsp", "DiasAplazEsp", "DiasTomEsp",
                                "DiasPenEsp", "ObsEsp", "EstadoEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodInte']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblApla']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodConc']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSecEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblSecApla']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecDisd']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFecDish']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDiaApla']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDiaToma']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDiaPend']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtObsErva_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtObsErvaSoli_lblTexto']"

                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Numero de Contrato", "Identificación", "Nombres y Apellidos", "Cod. Interno",
                                        "Aplazamiento", "Concepto Asociado", "Secuencial", "Aplazamiento 2", "Fecha Inicial",
                                        "Fecha Final", "Dias Aplazados", "Dias Tomados", "Dias Pendientes", "Observaciones",
                                        "Observaciones de solicitud"

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
                                        if (counter == 13) for (int i = 0; i < 3; i++) SendKeys.Send("{DOWN}");
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
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
        public void SmartPeople_frmNmConMarcLNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmNmConMarcLNTC")
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
                                "HomeEsp", "FiltroEsp", "FechaIniEsp", "FechaFinEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblFiltro']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaIni_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFechaFin_lblFecha']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Filtro", "Fecha Inicial", "Fecha Final"
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
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
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
        public void SmartPeople_frmNmConviaNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmNmConviaNTC")
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
                                "HomeEsp", "NumContEsp", "IdentEsp", "NomEsp", "ApellEsp", "TelefEsp", "TelefMovEsp", "N1Esp",
                                "N2Esp", "N3Esp", "N4Esp", "N5Esp", "N6Esp", "N7Esp", "DepenEsp", "TipoViaEsp", "MotCoEsp", "DescTraEsp",
                                "ObsSolEsp", "NumResEsp", "FechaResEsp", "DiasPerEsp", "DiasNoPerEsp", "OrDesEsp", "TotalViDiEsp",
                                "TotalViTeEsp", "ToViEsp", "ViatiDoEsp", "TotalDoEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblModAcad']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label4']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label5']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label6']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTelResi']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTelMovi0']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNivel1']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNivel2']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNivel3']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNivel4']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNivel5']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNivel6']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNivel7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDepend10']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTipEsca']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMotViat']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtDesComi_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlTxtObser_lblTexto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTelResi0']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecReso_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDiaPern']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblDiaNper']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblMulDest']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblViaDiar']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTraTerr']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTotViat']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblViaDola']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblTotDola']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Numero de Contrato", "Identificación", "Nombre", "Apellido", "Teléfono",
                                        "Teléfono Movil", "Nivel 1", "Nivel 2", "Nivel 3", "Nivel 4", "Nivel 5",
                                        "Nivel 6", "Nivel 7", "Dependencia", "Tipo de Viático", "Motivo de Comisión",
                                        "Descripción del Trabajo o Capacitación", "Observaciones de la Solicitud",
                                        "Número de Resolución", "Fecha Resolución", "Días Pernoctados",
                                        "Días No Pernoctados", "Origen y Destino", "Total Viáticos Diarios",
                                        "Total Viáticos Terrestres", "Total Viáticos", "Viáticos Diarios Dólares",
                                        "Total Dólares"
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
                                        if (counter == 17) for (int i = 0; i < 12; i++) SendKeys.Send("{DOWN}");
                                        if (counter == 22) for (int i = 0; i < 6; i++) SendKeys.Send("{DOWN}");
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
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$txt')]")));
                                    //elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));
                                    if (elementList.Count > 0)
                                    {
                                        elementList[5].Click();
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
                                    selenium.Screenshot("Error, No permite TABS en los campos", true, file);

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
        public void SmartPeople_frmNmDviclNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmNmDviclNTC")
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
                                "HomeEsp", "IdentEsp", "NomEsp", "NumContEsp", "FechaSolEsp", "ConceptEsp", "ValEsp",
                                "ValorIvaEsp", "ValorAntiEsp", "ValTotalAntiEsp"
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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblNoCont']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFecViat_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblConcepto']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValViat']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValIvav']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValorAnti']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_lblValorViatTRM']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Identificación", "Nombre", "Nro de Contrato", "Fecha de Solicitud", "Concepto",
                                        "Valor", "Valor Iva", "Valor Anticipo", "Valor total Anticipo en USD"
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
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$txt')]")));
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
        public void SmartPeople_frmNmNotcajNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.SmartPeople_NTC_2.SmartPeople_frmNmNotcajNTC")
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
                                "HomeEsp", "GuardarEsp", "IdentEsp", "NomApelEsp", "NumContEsp", "SucurEsp", "ConAperCajaEsp",
                                "NomCajaEsp", "CuantiaEsp", "ManejaCuRuEsp", "PerteEmpreEsp", "FechaTasaEsp", "IncluyeImpEsp",
                                "ManejaVarSuEsp", "ManejaToCuEsp", "NumDisEsp", "ResCentCostEsp", "TipoCajaEsp", "ValidarPagoEsp",
                                "CuentaContableEsp", "ObsSolEsp"
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


                                    List<string> name = new List<string>() { "Titulo", "Subtitulo", "Home", "Guardar" };
                                    string Titulo = selenium.Title();
                                    string Subtitulo = selenium.Subtitulo("ctl00_lblTitulo");
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    string Guardar = selenium.EmergenteBotones("btnGuardar");
                                    List<string> nameBotEn = new List<string>() { Titulo, Subtitulo, Home, Guardar };

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
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label7']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label12']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label15']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label8']", //Cuantia
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label14']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label17']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_KCtrlFectasa_lblFecha']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label18']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label4']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label20']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label23']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label24']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label25']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label28']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label29']",
                                        "xpath=//span[@id='ctl00_ContenidoPagina_Label5']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Identificación", "Nombres y Apellidos", "Nro Contrato", "Sucursal",
                                        "Controla Apertura de Caja", "Nombre de la Caja", "Cuantia", "Maneja Cuantia por Rubro",
                                        "Pertenece a Empresa", "Fecha de Tasa", "Incluye Impuestos para PG", "Maneja varias sucursales",
                                        "Maneja tope de cuantia", "Número Disponible", "Restringe por Centro de costo", "Tipo de Caja",
                                        "Validar pagos provisionales", "Cuenta contable", "Observaciones Solicitud"
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
                                        if (counter == 15) for (int i = 0; i < 6; i++) SendKeys.Send("{DOWN}");
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
