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
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using System.Data;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using APITest;
using Keys = OpenQA.Selenium.Keys;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;

namespace Web_Kactus_Test_V2
{
    [TestClass]
    public class Self_Service_FGN : FuncionesVitales
    {

        string Modulo = "FGN";
        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();

        public Self_Service_FGN()
        {
        }

        [TestMethod]
        public void NM_EnvíoCorreoVacacionesPrimerAprobadorConFlujoAreaRiesgoCorreoEstandar()
        {

            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Self_Service_FGN.NM_EnvíoCorreoVacacionesPrimerAprobadorConFlujoAreaRiesgoCorreoEstandar")
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
                                //Datos Mis Vacaciones    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string url = rows["url"].ToString();
                                //limpiar procesos
                                Process[] processes = Process.GetProcessesByName("chromedriver");
                                if (processes.Length > 0)
                                {
                                    for (int i = 0; i < processes.Length; i++)
                                    {
                                        processes[i].Kill();
                                    }
                                }
                                try
                                {
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);


                                    //eliminar registros previos

                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarVacaciones = $"Delete from NM_PROVA where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarVacaciones, database, user);

                                    //MIS SOLICITUDES 
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    //MIS VACACIONES
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Mis Vacaciones')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Mis Vacaciones", true, file);
                                    //NUEVA
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[2]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nueva Solicitud de Vacaciones", true, file);
                                    //FECHA INICIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']", FechaInicial);
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");

                                    //FECHA FINAL
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis2_txtFecha']", FechaFinal);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    //OBSERVACIONES
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(5000);
                                    //OBSERVACIONES
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']");
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    //APLICAR
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[3]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Vacaciones registrado", true, file);
                                    Thread.Sleep(7000);
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(15000);
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    ////////////////////////////////////////////////////
                                    //limpiar procesos
                                    Process[] processes1 = Process.GetProcessesByName("chromedriver");
                                    if (processes1.Length > 0)
                                    {
                                        for (int i = 0; i < processes1.Length; i++)
                                        {
                                            processes1[i].Kill();
                                        }
                                    }
                                    if (errorsTest.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorsTest);

                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}",
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
                                    Thread.Sleep(500);
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
        public void NM_EnvíoCorreoVacacionesPrimerAprobadorConFlujoAreaRiesgoCorreoKGnInfno()
        {

            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Self_Service_FGN.NM_EnvíoCorreoVacacionesPrimerAprobadorConFlujoAreaRiesgoCorreoKGnInfno")
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
     

                          
                            if (

                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                //Datos Mis Vacaciones    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string url = rows["url"].ToString();
                                //limpiar procesos
                                Process[] processes = Process.GetProcessesByName("chromedriver");
                                if (processes.Length > 0)
                                {
                                    for (int i = 0; i < processes.Length; i++)
                                    {
                                        processes[i].Kill();
                                    }
                                }
                                try
                                {
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);


                                    ////eliminar registros previos

                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarVacaciones = $"Delete from NM_PROVA where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarVacaciones, database, user);

                                    //MIS SOLICITUDES 
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    //MIS VACACIONES
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Mis Vacaciones')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Mis Vacaciones", true, file);
                                    //NUEVA
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[2]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nueva Solicitud de Vacaciones", true, file);
                                    //FECHA INICIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']", FechaInicial);
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");

                                    //FECHA FINAL
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis2_txtFecha']", FechaFinal);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    //OBSERVACIONES
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(5000);
                                    //OBSERVACIONES
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']");
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    //APLICAR
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[3]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Vacaciones registrado", true, file);
                                    Thread.Sleep(7000);
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(15000);
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    ////////////////////////////////////////////////////
                                    //limpiar procesos
                                    Process[] processes1 = Process.GetProcessesByName("chromedriver");
                                    if (processes1.Length > 0)
                                    {
                                        for (int i = 0; i < processes1.Length; i++)
                                        {
                                            processes1[i].Kill();
                                        }
                                    }
                                    if (errorsTest.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorsTest);

                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}",
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
                                    Thread.Sleep(500);
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
        public void NM_EnvíoCorreoVacacionesPrimerAprobadorConFlujoJefeInmediatoCorreoEstandar()
        {

            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Self_Service_FGN.NM_EnvíoCorreoVacacionesPrimerAprobadorConFlujoJefeInmediatoCorreoEstandar")
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
                                //Datos Mis Vacaciones    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string url = rows["url"].ToString();
                                //limpiar procesos
                                Process[] processes = Process.GetProcessesByName("chromedriver");
                                if (processes.Length > 0)
                                {
                                    for (int i = 0; i < processes.Length; i++)
                                    {
                                        processes[i].Kill();
                                    }
                                }
                                try
                                {
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);


                                    //eliminar registros previos

                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarVacaciones = $"Delete from NM_PROVA where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarVacaciones, database, user);

                                    //MIS SOLICITUDES 
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    //MIS VACACIONES
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Mis Vacaciones')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Mis Vacaciones", true, file);
                                    //NUEVA
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[2]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nueva Solicitud de Vacaciones", true, file);
                                    //FECHA INICIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']", FechaInicial);
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");

                                    //FECHA FINAL
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis2_txtFecha']", FechaFinal);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    //OBSERVACIONES
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(5000);
                                    //OBSERVACIONES
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']");
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    //APLICAR
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[3]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Vacaciones registrado", true, file);
                                    Thread.Sleep(7000);
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(15000);
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    ////////////////////////////////////////////////////
                                    //limpiar procesos
                                    Process[] processes1 = Process.GetProcessesByName("chromedriver");
                                    if (processes1.Length > 0)
                                    {
                                        for (int i = 0; i < processes1.Length; i++)
                                        {
                                            processes1[i].Kill();
                                        }
                                    }
                                    if (errorsTest.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorsTest);

                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}",
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
                                    Thread.Sleep(500);
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
        public void NM_EnvíoCorreoVacacionesPrimerAprobadorConFlujoJefeInmediatoCorreoKGnInfno()
        {

            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Self_Service_FGN.NM_EnvíoCorreoVacacionesPrimerAprobadorConFlujoJefeInmediatoCorreoKGnInfno")
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
                                //Datos Mis Vacaciones    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string url = rows["url"].ToString();
                                //limpiar procesos
                                Process[] processes = Process.GetProcessesByName("chromedriver");
                                if (processes.Length > 0)
                                {
                                    for (int i = 0; i < processes.Length; i++)
                                    {
                                        processes[i].Kill();
                                    }
                                }
                                try
                                {
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);


                                    //eliminar registros previos

                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarVacaciones = $"Delete from NM_PROVA where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarVacaciones, database, user);

                                    //MIS SOLICITUDES 
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    //MIS VACACIONES
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Mis Vacaciones')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Mis Vacaciones", true, file);
                                    //NUEVA
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[2]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nueva Solicitud de Vacaciones", true, file);
                                    //FECHA INICIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']", FechaInicial);
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");

                                    //FECHA FINAL
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis2_txtFecha']", FechaFinal);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    //OBSERVACIONES
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(5000);
                                    //OBSERVACIONES
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']");
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    //APLICAR
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[3]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Vacaciones registrado", true, file);
                                    Thread.Sleep(7000);
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(15000);
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    ////////////////////////////////////////////////////
                                    //limpiar procesos
                                    Process[] processes1 = Process.GetProcessesByName("chromedriver");
                                    if (processes1.Length > 0)
                                    {
                                        for (int i = 0; i < processes1.Length; i++)
                                        {
                                            processes1[i].Kill();
                                        }
                                    }
                                    if (errorsTest.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorsTest);

                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}",
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
                                    Thread.Sleep(500);
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
        public void NM_EnvíoCorreoVacacionesPrimerAprobadorJefeInmediatoSegundoAprobadorJefedelJefeCopiaJefeInmediatoAprobacionMasiva()
        {

            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Self_Service_FGN.NM_EnvíoCorreoVacacionesPrimerAprobadorJefeInmediatoSegundoAprobadorJefedelJefeCopiaJefeInmediatoAprobacionMasiva")
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
                                //Datos Mis Vacaciones    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();
                                string JefeJefe = rows["JefeJefe"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string url = rows["url"].ToString();
                                //limpiar procesos
                                Process[] processes = Process.GetProcessesByName("chromedriver");
                                if (processes.Length > 0)
                                {
                                    for (int i = 0; i < processes.Length; i++)
                                    {
                                        processes[i].Kill();
                                    }
                                }
                                try
                                {
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);


                                    //eliminar registros previos

                                    string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeJefe}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarSolicitud2 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud2, database, user);

                                    string eliminarVacaciones = $"Delete from NM_PROVA where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarVacaciones, database, user);

                                    //MIS SOLICITUDES 
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    //MIS VACACIONES
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Mis Vacaciones')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Mis Vacaciones", true, file);
                                    //NUEVA
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[2]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nueva Solicitud de Vacaciones", true, file);
                                    //FECHA INICIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']", FechaInicial);
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");

                                    //FECHA FINAL
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis2_txtFecha']", FechaFinal);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    //OBSERVACIONES
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(5000);
                                    //OBSERVACIONES
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']");
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    //APLICAR
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[3]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Vacaciones registrado", true, file);
                                    Thread.Sleep(7000);
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(15000);
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    //-------------------------------APROBAR REGISTRO-------------------------------------------------
                                    //Ingreso jefe1
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Jefe1", true, file);

                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//button[@id='pLider']");
                                    }
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Jefe Lider", true, file);

                                    selenium.Click("//a[contains(.,'MIS COLABORADORES')]");
                                    selenium.Click("//a[contains(.,'Vacaciones')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Vacaciones de Mis Colaboradores", true, file);

                                    if (selenium.ExistControl("//*[@id='tableVacaciones']/tbody[1]/tr/td[10]/a[1]/i[1]"))
                                    {
                                        selenium.Click("//*[@id='ApruebaMasivo']");
                                        Thread.Sleep(5000);
                                        selenium.Screenshot("Solicitud Aprobar Masiva", true, file);
                                        selenium.Click("//*[@id='tableVacacionesMas']/tbody/tr/td[2]/input");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Solicitud a Aprobar", true, file);
                                        selenium.Click("//*[@id='ctl00_ContentPopapModel_btnProcesarTodos']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Aprobada", true, file);
                                    }
                                    else
                                    {
                                        Assert.Fail("NO EXISTEN SOLICITUDES PARA APROBAR DE MIS COLABORADORES");
                                    }

                                    Thread.Sleep(1000);
                                    selenium.Close();
                                    ////////////////////////////////////////////////////
                                    //limpiar procesos
                                    Process[] processes1 = Process.GetProcessesByName("chromedriver");
                                    if (processes1.Length > 0)
                                    {
                                        for (int i = 0; i < processes1.Length; i++)
                                        {
                                            processes1[i].Kill();
                                        }
                                    }
                                    if (errorsTest.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorsTest);

                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}",
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
                                    Thread.Sleep(500);
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
        public void NM_EnvíoCorreoVacacionesPrimerAprobadorJefeInmediatoSegundoAprobadorAreaRiesgoCopiaJefeInmediatoNormalKGnInfno()
        {

            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Self_Service_FGN.NM_EnvíoCorreoVacacionesPrimerAprobadorJefeInmediatoSegundoAprobadorAreaRiesgoCopiaJefeInmediatoNormalKGnInfno")
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
                                //Datos Mis Vacaciones    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();
                                string JefeJefe = rows["JefeJefe"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string url = rows["url"].ToString();
                                //limpiar procesos
                                Process[] processes = Process.GetProcessesByName("chromedriver");
                                if (processes.Length > 0)
                                {
                                    for (int i = 0; i < processes.Length; i++)
                                    {
                                        processes[i].Kill();
                                    }
                                }
                                try
                                {
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);


                                    //eliminar registros previos

                                    string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeJefe}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarSolicitud2 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud2, database, user);

                                    string eliminarVacaciones = $"Delete from NM_PROVA where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarVacaciones, database, user);

                                    //MIS SOLICITUDES 
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    //MIS VACACIONES
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Mis Vacaciones')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Mis Vacaciones", true, file);
                                    //NUEVA
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[2]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nueva Solicitud de Vacaciones", true, file);
                                    //FECHA INICIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']", FechaInicial);
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");

                                    //FECHA FINAL
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis2_txtFecha']", FechaFinal);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    //OBSERVACIONES
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(5000);
                                    //OBSERVACIONES
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']");
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    //APLICAR
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[3]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Vacaciones registrado", true, file);
                                    Thread.Sleep(7000);
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(15000);
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    //-------------------------------APROBAR REGISTRO-------------------------------------------------
                                    //Ingreso jefe1
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Jefe1", true, file);

                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//button[@id='pLider']");
                                    }
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Jefe Lider", true, file);

                                    selenium.Click("//a[contains(.,'MIS COLABORADORES')]");
                                    selenium.Click("//a[contains(.,'Vacaciones')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Vacaciones de Mis Colaboradores", true, file);

                                    if (selenium.ExistControl("//*[@id='tableVacaciones']/tbody[1]/tr/td[10]/a[1]/i[1]"))
                                    {
                                        selenium.ScrollTo("0", "300");
                                        Thread.Sleep(5000);
                                        selenium.Click("//*[@id='tableVacaciones']/tbody[1]/tr/td[10]/a[1]/i[1]");
                                        Thread.Sleep(5000);
                                        selenium.Screenshot("Solicitud Aprobar", true, file);
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_Aprueba']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Aprobada", true, file);
                                    }
                                    else
                                    {
                                        Assert.Fail("NO EXISTEN SOLICITUDES PARA APROBAR DE MIS COLABORADORES");
                                    }

                                    Thread.Sleep(1000);
                                    selenium.Close();
                                    ////////////////////////////////////////////////////
                                    //limpiar procesos
                                    Process[] processes1 = Process.GetProcessesByName("chromedriver");
                                    if (processes1.Length > 0)
                                    {
                                        for (int i = 0; i < processes1.Length; i++)
                                        {
                                            processes1[i].Kill();
                                        }
                                    }
                                    if (errorsTest.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorsTest);

                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}",
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
                                    Thread.Sleep(500);
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
        public void NM_EnvíoCorreoVacacionesPrimerAprobadorJefeInmediatoSegundoAprobadorAreaRiesgoCopiaJefeInmediatoAprobacionMasiva()
        {

            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Self_Service_FGN.NM_EnvíoCorreoVacacionesPrimerAprobadorJefeInmediatoSegundoAprobadorAreaRiesgoCopiaJefeInmediatoAprobacionMasiva")
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
                                //Datos Mis Vacaciones    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();
                                string JefeJefe = rows["JefeJefe"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string url = rows["url"].ToString();
                                //limpiar procesos
                                Process[] processes = Process.GetProcessesByName("chromedriver");
                                if (processes.Length > 0)
                                {
                                    for (int i = 0; i < processes.Length; i++)
                                    {
                                        processes[i].Kill();
                                    }
                                }
                                try
                                {
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);


                                    //eliminar registros previos

                                    string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeJefe}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarSolicitud2 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud2, database, user);

                                    string eliminarVacaciones = $"Delete from NM_PROVA where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarVacaciones, database, user);

                                    //MIS SOLICITUDES 
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    //MIS VACACIONES
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Mis Vacaciones')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Mis Vacaciones", true, file);
                                    //NUEVA
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[2]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nueva Solicitud de Vacaciones", true, file);
                                    //FECHA INICIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']", FechaInicial);
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");

                                    //FECHA FINAL
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis2_txtFecha']", FechaFinal);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    //OBSERVACIONES
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(5000);
                                    //OBSERVACIONES
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']");
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    //APLICAR
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[3]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Vacaciones registrado", true, file);
                                    Thread.Sleep(7000);
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(15000);
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    //-------------------------------APROBAR REGISTRO-------------------------------------------------
                                    //Ingreso jefe1
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Jefe1", true, file);

                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//button[@id='pLider']");
                                    }
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Jefe Lider", true, file);

                                    selenium.Click("//a[contains(.,'MIS COLABORADORES')]");
                                    selenium.Click("//a[contains(.,'Vacaciones')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Vacaciones de Mis Colaboradores", true, file);

                                    if (selenium.ExistControl("//*[@id='tableVacaciones']/tbody[1]/tr/td[10]/a[1]/i[1]"))
                                    {
                                        selenium.Click("//*[@id='ApruebaMasivo']");
                                        Thread.Sleep(5000);
                                        selenium.Screenshot("Solicitud Aprobar Masiva", true, file);
                                        selenium.Click("//*[@id='tableVacacionesMas']/tbody/tr/td[2]/input");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Solicitud a Aprobar", true, file);
                                        selenium.Click("//*[@id='ctl00_ContentPopapModel_btnProcesarTodos']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Aprobada", true, file);
                                    }
                                    else
                                    {
                                        Assert.Fail("NO EXISTEN SOLICITUDES PARA APROBAR DE MIS COLABORADORES");
                                    }

                                    Thread.Sleep(1000);
                                    selenium.Close();
                                    ////////////////////////////////////////////////////
                                    //limpiar procesos
                                    Process[] processes1 = Process.GetProcessesByName("chromedriver");
                                    if (processes1.Length > 0)
                                    {
                                        for (int i = 0; i < processes1.Length; i++)
                                        {
                                            processes1[i].Kill();
                                        }
                                    }
                                    if (errorsTest.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorsTest);

                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}",
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
                                    Thread.Sleep(500);
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
        public void NM_EnvíoCorreoVacacionesPrimerAprobadorJefeInmediatoSegundoAprobadorAreaRiesgoCopiaJefeInmediatoAprobacionMasivaCorreoMO()
        {

            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Self_Service_FGN.NM_EnvíoCorreoVacacionesPrimerAprobadorJefeInmediatoSegundoAprobadorAreaRiesgoCopiaJefeInmediatoAprobacionMasivaCorreoMO")
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
                                //Datos Mis Vacaciones    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();
                                string JefeJefe = rows["JefeJefe"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string url = rows["url"].ToString();
                                //limpiar procesos
                                Process[] processes = Process.GetProcessesByName("chromedriver");
                                if (processes.Length > 0)
                                {
                                    for (int i = 0; i < processes.Length; i++)
                                    {
                                        processes[i].Kill();
                                    }
                                }
                                try
                                {
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);


                                    //eliminar registros previos

                                    string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeJefe}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarSolicitud2 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud2, database, user);

                                    string eliminarVacaciones = $"Delete from NM_PROVA where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarVacaciones, database, user);

                                    //MIS SOLICITUDES 
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    //MIS VACACIONES
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Mis Vacaciones')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Mis Vacaciones", true, file);
                                    //NUEVA
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[2]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nueva Solicitud de Vacaciones", true, file);
                                    //FECHA INICIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']", FechaInicial);
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");

                                    //FECHA FINAL
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis2_txtFecha']", FechaFinal);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    //OBSERVACIONES
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(5000);
                                    //OBSERVACIONES
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']");
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    //APLICAR
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[3]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Vacaciones registrado", true, file);
                                    Thread.Sleep(7000);
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(15000);
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    //-------------------------------APROBAR REGISTRO-------------------------------------------------
                                    //Ingreso jefe1
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Jefe1", true, file);

                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//button[@id='pLider']");
                                    }
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Jefe Lider", true, file);

                                    selenium.Click("//a[contains(.,'MIS COLABORADORES')]");
                                    selenium.Click("//a[contains(.,'Vacaciones')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Vacaciones de Mis Colaboradores", true, file);

                                    if (selenium.ExistControl("//*[@id='tableVacaciones']/tbody[1]/tr/td[10]/a[1]/i[1]"))
                                    {
                                        selenium.Click("//*[@id='ApruebaMasivo']");
                                        Thread.Sleep(5000);
                                        selenium.Screenshot("Solicitud Aprobar Masiva", true, file);
                                        selenium.Click("//*[@id='tableVacacionesMas']/tbody/tr/td[2]/input");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Solicitud a Aprobar", true, file);
                                        selenium.Click("//*[@id='ctl00_ContentPopapModel_btnProcesarTodos']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Aprobada", true, file);
                                    }
                                    else
                                    {
                                        Assert.Fail("NO EXISTEN SOLICITUDES PARA APROBAR DE MIS COLABORADORES");
                                    }

                                    Thread.Sleep(1000);
                                    selenium.Close();
                                    ////////////////////////////////////////////////////
                                    //limpiar procesos
                                    Process[] processes1 = Process.GetProcessesByName("chromedriver");
                                    if (processes1.Length > 0)
                                    {
                                        for (int i = 0; i < processes1.Length; i++)
                                        {
                                            processes1[i].Kill();
                                        }
                                    }
                                    if (errorsTest.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorsTest);

                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}",
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
                                    Thread.Sleep(500);
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
        public void NM_EnvíoCorreoVacacionesPrimerAprobadorJefeInmediatoSegundoAprobadorJefeJdelJefeCopiaJefeInmediatoCorreoEstandar()
        {

            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Self_Service_FGN.NM_EnvíoCorreoVacacionesPrimerAprobadorJefeInmediatoSegundoAprobadorJefeJdelJefeCopiaJefeInmediatoCorreoEstandar")
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
                                //Datos Mis Vacaciones    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();
                                string JefeJefe = rows["JefeJefe"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string url = rows["url"].ToString();
                                //limpiar procesos
                                Process[] processes = Process.GetProcessesByName("chromedriver");
                                if (processes.Length > 0)
                                {
                                    for (int i = 0; i < processes.Length; i++)
                                    {
                                        processes[i].Kill();
                                    }
                                }
                                try
                                {
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);


                                    //eliminar registros previos

                                    string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeJefe}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarSolicitud2 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud2, database, user);

                                    string eliminarVacaciones = $"Delete from NM_PROVA where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarVacaciones, database, user);

                                    //MIS SOLICITUDES 
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    //MIS VACACIONES
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Mis Vacaciones')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Mis Vacaciones", true, file);
                                    //NUEVA
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[2]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nueva Solicitud de Vacaciones", true, file);
                                    //FECHA INICIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']", FechaInicial);
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");

                                    //FECHA FINAL
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis2_txtFecha']", FechaFinal);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    //OBSERVACIONES
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(5000);
                                    //OBSERVACIONES
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']");
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    //APLICAR
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[3]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Vacaciones registrado", true, file);
                                    Thread.Sleep(7000);
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(15000);
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    //-------------------------------APROBAR REGISTRO-------------------------------------------------
                                    //Ingreso jefe1
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Jefe1", true, file);

                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//button[@id='pLider']");
                                    }
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Jefe Lider", true, file);

                                    selenium.Click("//a[contains(.,'MIS COLABORADORES')]");
                                    selenium.Click("//a[contains(.,'Vacaciones')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Vacaciones de Mis Colaboradores", true, file);

                                    if (selenium.ExistControl("//*[@id='tableVacaciones']/tbody[1]/tr/td[10]/a[1]/i[1]"))
                                    {
                                        selenium.ScrollTo("0", "300");
                                        Thread.Sleep(5000);
                                        selenium.Click("//*[@id='tableVacaciones']/tbody[1]/tr/td[10]/a[1]/i[1]");
                                        Thread.Sleep(5000);
                                        selenium.Screenshot("Solicitud Aprobar", true, file);
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_Aprueba']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Aprobada", true, file);
                                    }
                                    else
                                    {
                                        Assert.Fail("NO EXISTEN SOLICITUDES PARA APROBAR DE MIS COLABORADORES");
                                    }

                                    Thread.Sleep(1000);
                                    selenium.Close();
                                    ////////////////////////////////////////////////////
                                    //limpiar procesos
                                    Process[] processes1 = Process.GetProcessesByName("chromedriver");
                                    if (processes1.Length > 0)
                                    {
                                        for (int i = 0; i < processes1.Length; i++)
                                        {
                                            processes1[i].Kill();
                                        }
                                    }
                                    if (errorsTest.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorsTest);

                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}",
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
                                    Thread.Sleep(500);
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
        public void NM_EnvíoCorreoVacacionesPrimerAprobadorJefeInmediatoSegundoAprobadorJefeJdelJefeCopiaJefeInmediatoCorreoKGnInfno()
        {

            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Self_Service_FGN.NM_EnvíoCorreoVacacionesPrimerAprobadorJefeInmediatoSegundoAprobadorJefeJdelJefeCopiaJefeInmediatoCorreoKGnInfno")
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
                                //Datos Mis Vacaciones    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();
                                string JefeJefe = rows["JefeJefe"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string url = rows["url"].ToString();
                                //limpiar procesos
                                Process[] processes = Process.GetProcessesByName("chromedriver");
                                if (processes.Length > 0)
                                {
                                    for (int i = 0; i < processes.Length; i++)
                                    {
                                        processes[i].Kill();
                                    }
                                }
                                try
                                {
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);


                                    //eliminar registros previos

                                    string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeJefe}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarSolicitud2 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud2, database, user);

                                    string eliminarVacaciones = $"Delete from NM_PROVA where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarVacaciones, database, user);

                                    //MIS SOLICITUDES 
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    //MIS VACACIONES
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Mis Vacaciones')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Mis Vacaciones", true, file);
                                    //NUEVA
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[2]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nueva Solicitud de Vacaciones", true, file);
                                    //FECHA INICIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']", FechaInicial);
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");

                                    //FECHA FINAL
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis2_txtFecha']", FechaFinal);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    //OBSERVACIONES
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(5000);
                                    //OBSERVACIONES
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']");
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    //APLICAR
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[3]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Vacaciones registrado", true, file);
                                    Thread.Sleep(7000);
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(15000);
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    //-------------------------------APROBAR REGISTRO-------------------------------------------------
                                    //Ingreso jefe1
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Jefe1", true, file);

                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//button[@id='pLider']");
                                    }
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Jefe Lider", true, file);

                                    selenium.Click("//a[contains(.,'MIS COLABORADORES')]");
                                    selenium.Click("//a[contains(.,'Vacaciones')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Vacaciones de Mis Colaboradores", true, file);

                                    if (selenium.ExistControl("//*[@id='tableVacaciones']/tbody[1]/tr/td[10]/a[1]/i[1]"))
                                    {
                                        selenium.ScrollTo("0", "300");
                                        Thread.Sleep(5000);
                                        selenium.Click("//*[@id='tableVacaciones']/tbody[1]/tr/td[10]/a[1]/i[1]");
                                        Thread.Sleep(5000);
                                        selenium.Screenshot("Solicitud Aprobar", true, file);
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_Aprueba']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Aprobada", true, file);
                                    }
                                    else
                                    {
                                        Assert.Fail("NO EXISTEN SOLICITUDES PARA APROBAR DE MIS COLABORADORES");
                                    }

                                    Thread.Sleep(1000);
                                    selenium.Close();
                                    ////////////////////////////////////////////////////
                                    //limpiar procesos
                                    Process[] processes1 = Process.GetProcessesByName("chromedriver");
                                    if (processes1.Length > 0)
                                    {
                                        for (int i = 0; i < processes1.Length; i++)
                                        {
                                            processes1[i].Kill();
                                        }
                                    }
                                    if (errorsTest.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorsTest);

                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}",
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
                                    Thread.Sleep(500);
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
        public void NM_AprobacionVacacionesValidacionEncargo()
        {

            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Self_Service_FGN.NM_AprobacionVacacionesValidacionEncargo")
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
                                //Datos Mis Vacaciones    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();
                                string JefeJefe = rows["JefeJefe"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string url = rows["url"].ToString();
                                //limpiar procesos
                                Process[] processes = Process.GetProcessesByName("chromedriver");
                                if (processes.Length > 0)
                                {
                                    for (int i = 0; i < processes.Length; i++)
                                    {
                                        processes[i].Kill();
                                    }
                                }
                                try
                                {
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);


                                    //eliminar registros previos
                                    string encargo = $"update GN_DIGIF set VAL_VARI='S' where COD_VARI = 'K0000008' AND EMP_VARI = '{CodEmpre}'";
                                    db.UpdateDeleteInsert(encargo, database, user);

                                    string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeJefe}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{JefeUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarSolicitud2 = $"Delete from NM_SOLTR where tip_apli ='A' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud2, database, user);

                                    string eliminarVacaciones = $"Delete from NM_PROVA where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarVacaciones, database, user);

                                    //MIS SOLICITUDES 
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    //MIS VACACIONES
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Mis Vacaciones')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Mis Vacaciones", true, file);
                                    //NUEVA
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[2]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nueva Solicitud de Vacaciones", true, file);
                                    //FECHA INICIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis1_txtFecha']", FechaInicial);
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");

                                    //FECHA FINAL
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecDis2_txtFecha']", FechaFinal);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    //OBSERVACIONES
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(5000);
                                    //OBSERVACIONES
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']");
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KtxtObserSoli_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    //APLICAR
                                    selenium.Click("//div[@id='ctl00_pBotones']/div/a[3]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Vacaciones registrado", true, file);
                                    Thread.Sleep(7000);
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(15000);
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    //-------------------------------APROBAR REGISTRO-------------------------------------------------
                                    //Ingreso jefe1
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Jefe1", true, file);

                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//button[@id='pLider']");
                                    }
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Jefe Lider", true, file);

                                    selenium.Click("//a[contains(.,'MIS COLABORADORES')]");
                                    selenium.Click("//a[contains(.,'Vacaciones')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Vacaciones de Mis Colaboradores", true, file);

                                    if (selenium.ExistControl("//*[@id='tableVacaciones']/tbody[1]/tr/td[10]/a[1]/i[1]"))
                                    {
                                        selenium.ScrollTo("0", "300");
                                        Thread.Sleep(5000);
                                        selenium.Click("//*[@id='tableVacaciones']/tbody[1]/tr/td[10]/a[1]/i[1]");
                                        Thread.Sleep(5000);
                                        selenium.Screenshot("Encargo", true, file);
                                        
                                    }
                                    else
                                    {
                                        Assert.Fail("NO EXISTEN SOLICITUDES PARA APROBAR DE MIS COLABORADORES");
                                    }

                                    Thread.Sleep(1000);
                                    selenium.Close();
                                    Thread.Sleep(4000);
                                    //LimpiarProcesos();

                                    string encargo1 = $"update GN_DIGIF set VAL_VARI='N' where COD_VARI = 'K0000008' AND EMP_VARI = '{CodEmpre}'";
                                    db.UpdateDeleteInsert(encargo1, database, user);
                                    ////////////////////////////////////////////////////
                                    //limpiar procesos
                                    Process[] processes1 = Process.GetProcessesByName("chromedriver");
                                    if (processes1.Length > 0)
                                    {
                                        for (int i = 0; i < processes1.Length; i++)
                                        {
                                            processes1[i].Kill();
                                        }
                                    }
                                    if (errorsTest.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorsTest);

                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}",
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
                                    Thread.Sleep(500);
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
    }
}
