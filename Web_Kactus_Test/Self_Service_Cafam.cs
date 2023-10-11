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
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using System.Data;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Web_Kactus_Test.UIMapNuevoClasses;
using APITest;
using Keys = OpenQA.Selenium.Keys;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;

namespace Web_Kactus_Test
{
    [CodedUITest]
    public class Self_Service_Cafam : FuncionesVitales
    {

        string Modulo = "Cafam";
        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();

        public Self_Service_Cafam()
        {

        }

        [TestMethod]
        public void CambioClave()
        {

            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.CambioClave")
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

                            if (
                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["ClaveAnt"].ToString().Length != 0 && rows["ClaveAnt"].ToString() != null &&
                                rows["ClaveNueva"].ToString().Length != 0 && rows["ClaveNueva"].ToString() != null 

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string ClaveAnt= rows["ClaveAnt"].ToString();
                                string ClaveNueva= rows["ClaveNueva"].ToString();

                                try
                                {
                                    string database = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";

                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";

                                    }

                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";

                                    }
                                    if (url.ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";


                                    }

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A CAMBIO CLAVE
                                    selenium.Scroll("//a[contains(.,'CAMBIAR CLAVE')]");
                                    selenium.Click("//a[contains(.,'CAMBIAR CLAVE')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Cambiar Clave')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("CAMBIO DE CLAVE", true, file);
                                    Thread.Sleep(2000);
                                    //CLAVE ANTERIOR
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtPasAnte']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtPasAnte']", ClaveAnt);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("CLAVE ANTERIOR", true, file);
                                    Thread.Sleep(5000);
                                    //CLAVE NUEVA
                                    Keyboard.SendKeys("{TAB}");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtPasNuev']", ClaveNueva);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("CLAVE NUEVA", true, file);
                                    Thread.Sleep(5000);
                                    //CONFIRMACION CLAVE
                                    Keyboard.SendKeys("{TAB}");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtPasConf']", ClaveNueva);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("CONFIRMACION CLAVE", true, file);
                                    //ACEPTAR
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnCamClav']");
                                    Thread.Sleep(4000);
                                    Screenshot("CONFIRMACION CAMBIO", true, file);
                                    //ACEPTAR ALERTA
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    //INICIAR SESION NUEVA CLAVE
                                    selenium.LoginApps(app, EmpleadoUser, ClaveNueva, url, file);
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Login Nueva Clave", true, file);

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(5000);
                                    selenium.Close();

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
        public void NovedadesTemporalesAprobacionRRHH()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.NovedadesTemporalesAprobacionRRHH")
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

                            if (
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["JefeUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["JefePass"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                rows["Cantidad"].ToString().Length != 0 && rows["Cantidad"].ToString() != null &&
                                rows["ValCuota"].ToString().Length != 0 && rows["ValCuota"].ToString() != null &&
                                rows["NumCuotas"].ToString().Length != 0 && rows["NumCuotas"].ToString() != null &&
                                rows["ValTotal"].ToString().Length != 0 && rows["ValTotal"].ToString() != null &&
                                rows["CodConc"].ToString().Length != 0 && rows["CodConc"].ToString() != null &&
                                rows["TipApli"].ToString().Length != 0 && rows["TipApli"].ToString() != null &&
                                rows["IndActi"].ToString().Length != 0 && rows["IndActi"].ToString() != null &&
                                rows["IndActi2"].ToString().Length != 0 && rows["IndActi2"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Concepto = rows["Concepto"].ToString();
                                string Cantidad = rows["Cantidad"].ToString();
                                string ValCuota = rows["ValCuota"].ToString();
                                string NumCuotas = rows["NumCuotas"].ToString();
                                string ValTotal = rows["ValTotal"].ToString();
                                string CodConc = rows["CodConc"].ToString();
                                string TipApli = rows["TipApli"].ToString();
                                string IndActi = rows["IndActi"].ToString();
                                string IndActi2 = rows["IndActi2"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();

                                try
                                {
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //PARAMETRIZACION
                                    if (database == "SQL")
                                    {
                                        string eliminarRegistro = $"DELETE NM_NOVTE WHERE COD_EMPR = '{CodEmpr}' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro2 = $"DELETE NM_NOVED WHERE COD_EMPR = '{CodEmpr}' AND COD_EMPL = '{EmpleadoUser}' AND COD_CONC = '{CodConc}'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro3 = $"DELETE NM_SOLTR where COD_SOLI='232' AND TIP_APLI='T'";
                                        db.UpdateDeleteInsert(eliminarRegistro3, database, user);
                                    }
                                    else
                                    {
                                        string eliminarRegistro = $"DELETE NM_NOVTE WHERE COD_EMPR = '421' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro2 = $"DELETE NM_NOVED WHERE COD_EMPR = '421' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro3 = $"DELETE NM_SOLTR where COD_SOLI='49' AND TIP_APLI='T'";
                                        db.UpdateDeleteInsert(eliminarRegistro3, database, user);
                                    }
                                    //INICIO PRUEBA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'NOVEDADES')]");
                                    selenium.Click("//a[contains(.,'NOVEDADES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Novedades Temporales colaborador')]");
                                    selenium.Click("//a[contains(.,'Novedades Temporales colaborador')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Novedades Temporales colaborador", true, file);

                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomConc']", Concepto);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtCanNove']", Cantidad);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtValCuot']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValCuot']", ValCuota);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNumCuot']", NumCuotas);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtValTota']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValTota']", ValTotal);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValTota']", ValTotal);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Datos Ingresados", true, file);

                                    selenium.Click("//a[@id='btnGuardar']");

                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Novedad Registrada", true, file);
                                  
                                    Thread.Sleep(5000);
                                    selenium.Close();

                                    //RRHH

                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                        Thread.Sleep(2000);
                                        if (database == "ORA")
                                        {
                                            selenium.Click("//button[contains(.,'GESTION HUMANA')]");
                                        }
                                        else
                                        {
                                            selenium.Click("//button[contains(.,'Rol RRHH')]");
                                        }
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Rol RRHH", true, file);

                                        selenium.Scroll("//a[contains(.,'PUESTO DE TRABAJO')]");
                                        selenium.Click("//a[contains(.,'PUESTO DE TRABAJO')]");
                                        Thread.Sleep(2000);
                                        selenium.Scroll("//a[contains(.,'Novedades Temporales RRHH')]");
                                        selenium.Click("//a[contains(.,'Novedades Temporales RRHH')]");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Novedades Temporales RRHH", true, file);

                                        if (database == "ORA")
                                        {
                                            selenium.Click("//div[@id='ctl00_pBotones']/div");
                                            selenium.Scroll("//*[@id='tablaSolicitudes']/tbody/tr/td[10]/a/i");
                                            Thread.Sleep(1000);
                                            selenium.Click("//*[@id='tablaSolicitudes']/tbody/tr/td[10]/a/i");

                                        }
                                        else
                                        {
                                            selenium.Scroll("//*[@id='tablaSolicitudes']/tbody/tr/td[10]/a/i");
                                            selenium.Click("//*[@id='tablaSolicitudes']/tbody/tr/td[10]/a/i");
                                        }
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Novedad", true, file);

                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgNmNovte_ctl02_LinkButton1']/i");
                                        Thread.Sleep(4000);
                                        selenium.Screenshot("Detalle solicitud", true, file);

                                        Thread.Sleep(2000);
                                    selenium.Screenshot("Solicitud Aprobada", true, file);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_Aprueba')]");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Envio Correo", true, file);
                                    Thread.Sleep(5000);
                                    selenium.Close();


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
        public void NovedadesTemporalesAprobacionMasivaRRHH()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.NovedadesTemporalesAprobacionMasivaRRHH")
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

                            if (
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["JefeUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["JefePass"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                rows["Cantidad"].ToString().Length != 0 && rows["Cantidad"].ToString() != null &&
                                rows["ValCuota"].ToString().Length != 0 && rows["ValCuota"].ToString() != null &&
                                rows["NumCuotas"].ToString().Length != 0 && rows["NumCuotas"].ToString() != null &&
                                rows["ValTotal"].ToString().Length != 0 && rows["ValTotal"].ToString() != null &&
                                rows["CodConc"].ToString().Length != 0 && rows["CodConc"].ToString() != null &&
                                rows["TipApli"].ToString().Length != 0 && rows["TipApli"].ToString() != null &&
                                rows["IndActi"].ToString().Length != 0 && rows["IndActi"].ToString() != null &&
                                rows["IndActi2"].ToString().Length != 0 && rows["IndActi2"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Concepto = rows["Concepto"].ToString();
                                string Cantidad = rows["Cantidad"].ToString();
                                string ValCuota = rows["ValCuota"].ToString();
                                string NumCuotas = rows["NumCuotas"].ToString();
                                string ValTotal = rows["ValTotal"].ToString();
                                string CodConc = rows["CodConc"].ToString();
                                string TipApli = rows["TipApli"].ToString();
                                string IndActi = rows["IndActi"].ToString();
                                string IndActi2 = rows["IndActi2"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();

                                try
                                {
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //PARAMETRIZACION
                                    if (database == "SQL")
                                    {
                                        string eliminarRegistro = $"DELETE NM_NOVTE WHERE COD_EMPR = '{CodEmpr}' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro2 = $"DELETE NM_NOVED WHERE COD_EMPR = '{CodEmpr}' AND COD_EMPL = '{EmpleadoUser}' AND COD_CONC = '{CodConc}'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro3 = $"DELETE NM_SOLTR where COD_SOLI='232' AND TIP_APLI='T'";
                                        db.UpdateDeleteInsert(eliminarRegistro3, database, user);
                                    }else
                                    {
                                        string eliminarRegistro = $"DELETE NM_NOVTE WHERE COD_EMPR = '421' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro2 = $"DELETE NM_NOVED WHERE COD_EMPR = '421' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro3 = $"DELETE NM_SOLTR where COD_SOLI='49' AND TIP_APLI='T'";
                                        db.UpdateDeleteInsert(eliminarRegistro3, database, user);
                                    }
                                    //INICIO PRUEBA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'NOVEDADES')]");
                                    selenium.Click("//a[contains(.,'NOVEDADES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Novedades Temporales colaborador')]");
                                    selenium.Click("//a[contains(.,'Novedades Temporales colaborador')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Novedades Temporales colaborador", true, file);

                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomConc']", Concepto);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtCanNove']", Cantidad);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValCuot']", ValCuota);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNumCuot']", NumCuotas);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtValTota']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValTota']", ValTotal);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValTota']", ValTotal);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Datos Ingresados", true, file);

                                    selenium.Click("//a[@id='btnGuardar']");

                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Novedad Registrada", true, file);

                                    
                                    Thread.Sleep(5000);
                                    selenium.Close();

                                    //RRHH

                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "ORA")
                                    {
                                        selenium.Click("//button[contains(.,'GESTION HUMANA')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//button[contains(.,'Rol RRHH')]");
                                    }
                                    Thread.Sleep(2000);

                                    selenium.Scroll("//a[contains(.,'PUESTO DE TRABAJO')]");
                                    selenium.Click("//a[contains(.,'PUESTO DE TRABAJO')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Novedades Temporales RRHH')]");
                                    selenium.Click("//a[contains(.,'Novedades Temporales RRHH')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Novedades Temporales RRHH", true, file);

                                    //APREOBACION MASIVA
                                    selenium.Click("//button[@id='btnAprobacionMasiva']");
                                    Thread.Sleep(3000);

                                    //SELECCIONAR
                                    selenium.Click("(//input[@type='checkbox'])[5]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("APROBACION MASIVA", true, file);
                                    Thread.Sleep(2000);
                                    //PROESAR
                                    selenium.Click("//input[@id='ctl00_ContentPopapModel_btnProcesarTodos']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Solicitud Aprobada",true,file);
                                    Thread.Sleep(5000);

                                    selenium.Close();


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
        public void NovedadesTemporalesExportarRRHH()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.NovedadesTemporalesExportarRRHH")
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

                            if (
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["JefeUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["JefePass"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                rows["Cantidad"].ToString().Length != 0 && rows["Cantidad"].ToString() != null &&
                                rows["ValCuota"].ToString().Length != 0 && rows["ValCuota"].ToString() != null &&
                                rows["NumCuotas"].ToString().Length != 0 && rows["NumCuotas"].ToString() != null &&
                                rows["ValTotal"].ToString().Length != 0 && rows["ValTotal"].ToString() != null &&
                                rows["CodConc"].ToString().Length != 0 && rows["CodConc"].ToString() != null &&
                                rows["TipApli"].ToString().Length != 0 && rows["TipApli"].ToString() != null &&
                                rows["IndActi"].ToString().Length != 0 && rows["IndActi"].ToString() != null &&
                                rows["IndActi2"].ToString().Length != 0 && rows["IndActi2"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Concepto = rows["Concepto"].ToString();
                                string Cantidad = rows["Cantidad"].ToString();
                                string ValCuota = rows["ValCuota"].ToString();
                                string NumCuotas = rows["NumCuotas"].ToString();
                                string ValTotal = rows["ValTotal"].ToString();
                                string CodConc = rows["CodConc"].ToString();
                                string TipApli = rows["TipApli"].ToString();
                                string IndActi = rows["IndActi"].ToString();
                                string IndActi2 = rows["IndActi2"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();

                                try
                                {
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //PARAMETRIZACION
                                    if (database == "SQL")
                                    {
                                        string eliminarRegistro = $"DELETE NM_NOVTE WHERE COD_EMPR = '{CodEmpr}' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro2 = $"DELETE NM_NOVED WHERE COD_EMPR = '{CodEmpr}' AND COD_EMPL = '{EmpleadoUser}' AND COD_CONC = '{CodConc}'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro3 = $"DELETE NM_SOLTR where COD_RESP='{EmpleadoUser}' AND TIP_APLI='T'";
                                        db.UpdateDeleteInsert(eliminarRegistro3, database, user);
                                    }
                                    else
                                    {
                                        string eliminarRegistro = $"DELETE NM_NOVTE WHERE COD_EMPR = '421' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro2 = $"DELETE NM_NOVED WHERE COD_EMPR = '421' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro3 = $"DELETE NM_SOLTR where COD_RESP='{EmpleadoUser}' AND TIP_APLI='T'";
                                        db.UpdateDeleteInsert(eliminarRegistro3, database, user);
                                    }

                                    //INICIO PRUEBA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'NOVEDADES')]");
                                    selenium.Click("//a[contains(.,'NOVEDADES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Novedades Temporales colaborador')]");
                                    selenium.Click("//a[contains(.,'Novedades Temporales colaborador')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Novedades Temporales colaborador", true, file);

                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomConc']", Concepto);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtCanNove']", Cantidad);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValCuot']", ValCuota);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNumCuot']", NumCuotas);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtValTota']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValTota']", ValTotal);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValTota']", ValTotal);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Datos Ingresados", true, file);

                                    selenium.Click("//a[@id='btnGuardar']");

                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Novedad Registrada", true, file);


                                    Thread.Sleep(5000);
                                    selenium.Close();

                                    //RRHH

                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "ORA")
                                    {
                                        selenium.Click("//button[contains(.,'GESTION HUMANA')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//button[contains(.,'Rol RRHH')]");
                                    }
                                    Thread.Sleep(2000);

                                    selenium.Scroll("//a[contains(.,'PUESTO DE TRABAJO')]");
                                    selenium.Click("//a[contains(.,'PUESTO DE TRABAJO')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Novedades Temporales RRHH')]");
                                    selenium.Click("//a[contains(.,'Novedades Temporales RRHH')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Novedades Temporales RRHH", true, file);

                                    //EXPORTAR
                                    selenium.Click("//div[@id='divExportar']/input");
                                    Thread.Sleep(5000);
                                    Screenshot("Excel Descargado", true, file);


                                    //Abrir Excel
                                    string excelPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ReporteProcesos.xls");
                                    Process.Start(excelPath);
                                    Thread.Sleep(60000);
                                    Screenshot("EXCEL ABIERTO", true, file);
                                    Thread.Sleep(20000);
                                    LimpiarProcesos();
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    
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
        public void NovedadesTemporalesFiltroRRHH()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.NovedadesTemporalesFiltroRRHH")
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

                            if (
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["JefeUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["JefePass"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                rows["Cantidad"].ToString().Length != 0 && rows["Cantidad"].ToString() != null &&
                                rows["ValCuota"].ToString().Length != 0 && rows["ValCuota"].ToString() != null &&
                                rows["NumCuotas"].ToString().Length != 0 && rows["NumCuotas"].ToString() != null &&
                                rows["ValTotal"].ToString().Length != 0 && rows["ValTotal"].ToString() != null &&
                                rows["CodConc"].ToString().Length != 0 && rows["CodConc"].ToString() != null &&
                                rows["TipApli"].ToString().Length != 0 && rows["TipApli"].ToString() != null &&
                                rows["IndActi"].ToString().Length != 0 && rows["IndActi"].ToString() != null &&
                                rows["IndActi2"].ToString().Length != 0 && rows["IndActi2"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Concepto = rows["Concepto"].ToString();
                                string Cantidad = rows["Cantidad"].ToString();
                                string ValCuota = rows["ValCuota"].ToString();
                                string NumCuotas = rows["NumCuotas"].ToString();
                                string ValTotal = rows["ValTotal"].ToString();
                                string CodConc = rows["CodConc"].ToString();
                                string TipApli = rows["TipApli"].ToString();
                                string IndActi = rows["IndActi"].ToString();
                                string IndActi2 = rows["IndActi2"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();
                                string FechaIni= rows["FechaIni"].ToString();
                                string FechaFin = rows["FechaFin"].ToString();
                                string FiltroTabla= rows["FiltroTabla"].ToString();
                                string ConceptoFiltro = rows["ConceptoFiltro"].ToString();


                                try
                                {
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //PARAMETRIZACION
                                    if (database == "SQL")
                                    {
                                        string eliminarRegistro = $"DELETE NM_NOVTE WHERE COD_EMPR = '{CodEmpr}' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro2 = $"DELETE NM_NOVED WHERE COD_EMPR = '{CodEmpr}' AND COD_EMPL = '{EmpleadoUser}' AND COD_CONC = '{CodConc}'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro3 = $"DELETE NM_SOLTR where COD_RESP='{EmpleadoUser}' AND TIP_APLI='T'";
                                        db.UpdateDeleteInsert(eliminarRegistro3, database, user);
                                    }
                                    else
                                    {
                                        string eliminarRegistro = $"DELETE NM_NOVTE WHERE COD_EMPR = '421' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro2 = $"DELETE NM_NOVED WHERE COD_EMPR = '421' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro3 = $"DELETE NM_SOLTR where COD_RESP='{EmpleadoUser}' AND TIP_APLI='T'";
                                        db.UpdateDeleteInsert(eliminarRegistro3, database, user);
                                    }

                                    //INICIO PRUEBA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'NOVEDADES')]");
                                    selenium.Click("//a[contains(.,'NOVEDADES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Novedades Temporales colaborador')]");
                                    selenium.Click("//a[contains(.,'Novedades Temporales colaborador')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Novedades Temporales colaborador", true, file);

                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomConc']", Concepto);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtCanNove']", Cantidad);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValCuot']", ValCuota);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNumCuot']", NumCuotas);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtValTota']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValTota']", ValTotal);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValTota']", ValTotal);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Datos Ingresados", true, file);

                                    selenium.Click("//a[@id='btnGuardar']");

                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Novedad Registrada", true, file);


                                    Thread.Sleep(5000);
                                    selenium.Close();

                                    //RRHH

                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "ORA")
                                    {
                                        selenium.Click("//button[contains(.,'GESTION HUMANA')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//button[contains(.,'Rol RRHH')]");
                                    }
                                    Thread.Sleep(2000);

                                    selenium.Scroll("//a[contains(.,'PUESTO DE TRABAJO')]");
                                    selenium.Click("//a[contains(.,'PUESTO DE TRABAJO')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Novedades Temporales RRHH')]");
                                    selenium.Click("//a[contains(.,'Novedades Temporales RRHH')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Novedades Temporales RRHH", true, file);

                                    //FILTRO CONCEPTO
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomConc']", ConceptoFiltro);
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Filtro por Concepto", true, file);
                                    //FILTRO FECHAS
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", FechaIni);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']", FechaFin);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Filtro por fechas", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Filtro por fechas", true, file);
                                    //FILTRO DATATABLE
                                    selenium.SendKeys("//div[@id='tablaSolicitudes_filter']/label/input", FiltroTabla);
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Filtro por datatable", true, file);

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
        public void ReporteNivelEndeudamientoPedWebcNominaDetallado()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.ReporteNivelEndeudamientoPedWebcNominaDetallado")
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

                            if (
                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();

                                try
                                {

                                    string database = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";

                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";

                                    }

                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";

                                    }
                                    if (url.ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";


                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Scroll("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Reportes", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(1500);

                                    //CHECK NIVEL ENDEUDAMIENTO
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(2000);
                                    //fechas
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlFecDesdA']", "2021");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);

                                    //TIPO DE REPORTE NOMINA
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblTipReport_0']");
                                    Thread.Sleep(2000);
                                    //FORMA REPORTE DETALLADO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblNivEnde_1']");
                                    Thread.Sleep(2000);

                                    selenium.Screenshot("Nomina Detallado", true, file);

                                    //GENERAR
                                    selenium.Click("//a[@id='ctl00_btnGenerar']");
                                    Thread.Sleep(8000);
                                    Screenshot("Reporte Endeudamiento Generado", true, file);
                                    Thread.Sleep(6000);

                                    //Abrir pdf
                                    if (database == "ORA")
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220510161359961.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);
                                    }
                                    else
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220509163526270.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);

                                    } 

                                    Screenshot("PDF ABIERTO", true, file);
                                    Thread.Sleep(20000);
                                    LimpiarProcesos();
                                    Thread.Sleep(2000);
                                    selenium.Close();

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
        public void ReporteNivelEndeudamientoPedWebcPrimaDetallado()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.ReporteNivelEndeudamientoPedWebcPrimaDetallado")
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

                            if (
                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();

                                try
                                {

                                    string database = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";

                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";

                                    }

                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";

                                    }
                                    if (url.ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";


                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Scroll("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Reportes", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(1500);
                                    //CHECK NIVEL ENDEUDAMIENTO
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(2000);
                                    //fechas
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlFecDesdA']", "2021");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    //TIPO DE REPORTE PRIMA
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblTipReport_1']");
                                    Thread.Sleep(2000);
                                    //FORMA REPORTE DETALLADO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblNivEnde_1']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Prima Detallado", true, file);
                                    Thread.Sleep(2000);

                                    //GENERAR
                                    selenium.Click("//a[@id='ctl00_btnGenerar']");
                                    Thread.Sleep(8000);
                                    Screenshot("Reporte Endeudamiento Generado", true, file);
                                    Thread.Sleep(6000);

                                    //Abrir pdf
                                    if (database == "ORA")
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220510161431644.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);
                                    }
                                    else
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220509163640902.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);

                                    }

                                    Screenshot("PDF ABIERTO", true, file);
                                    Thread.Sleep(20000);
                                    LimpiarProcesos();
                                    Thread.Sleep(2000);
                                    selenium.Close();

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
        public void ReporteNivelEndeudamientoPedWebcNominaResumido()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.ReporteNivelEndeudamientoPedWebcNominaResumido")
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

                            if (
                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();

                                try
                                {

                                    string database = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";

                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";

                                    }

                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";

                                    }
                                    if (url.ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";


                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Scroll("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Reportes", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(1500);
                                    //CHECK NIVEL ENDEUDAMIENTO
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(2000);
                                    //fechas
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlFecDesdA']", "2021");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);

                                    //TIPO DE REPORTE NOMINA
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblTipReport_0']");
                                    Thread.Sleep(2000);
                                    //FORMA REPORTE RESUMIDO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblNivEnde_0']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Nomina Resumido", true, file);

                                    //GENERAR
                                    selenium.Click("//a[@id='ctl00_btnGenerar']");
                                    Thread.Sleep(8000);
                                    Screenshot("Reporte Endeudamiento Generado", true, file);
                                    Thread.Sleep(6000);

                                    //Abrir pdf
                                    if (database == "ORA")
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220510161419438.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);
                                    }
                                    else
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220509163628965.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);

                                    }

                                    Screenshot("PDF ABIERTO", true, file);
                                    Thread.Sleep(20000);
                                    LimpiarProcesos();
                                    Thread.Sleep(2000);
                                    selenium.Close();

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
        public void ReporteNivelEndeudamientoPedWebcPrimaResumido()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.ReporteNivelEndeudamientoPedWebcPrimaResumido")
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

                            if (
                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();

                                try
                                {

                                    string database = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";

                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";

                                    }

                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";

                                    }
                                    if (url.ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";


                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Scroll("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Reportes", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(1500);
                                    //CHECK NIVEL ENDEUDAMIENTO
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(2000);
                                    //fechas
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlFecDesdA']", "2021");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    //TIPO DE REPORTE PRIMA
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblTipReport_1']");
                                    Thread.Sleep(2000);
                                    //FORMA REPORTE RESUMIDO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblNivEnde_0']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Prima Resumido", true, file);
                                    Thread.Sleep(2000);

                                    //GENERAR
                                    selenium.Click("//a[@id='ctl00_btnGenerar']");
                                    Thread.Sleep(8000);
                                    Screenshot("Reporte Endeudamiento Generado", true, file);
                                    Thread.Sleep(6000);

                                    //Abrir pdf
                                    if (database == "ORA")
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220510161442083.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);
                                    }
                                    else
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220509163652655.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);

                                    }

                                    Screenshot("PDF ABIERTO", true, file);
                                    Thread.Sleep(20000);
                                    LimpiarProcesos();
                                    Thread.Sleep(2000);
                                    selenium.Close();

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
        public void ReporteNivelEndeudamientoWebConfigNominaDetallado()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.ReporteNivelEndeudamientoWebConfigNominaDetallado")
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

                            if (
                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();

                                try
                                {

                                    string database = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";

                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";

                                    }

                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";

                                    }
                                    if (url.ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";


                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Scroll("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Reportes", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(1500);
                                    //CHECK NIVEL ENDEUDAMIENTO
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(2000);
                                    //fechas
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlFecDesdA']", "2021");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    //TIPO DE REPORTE NOMINA
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblTipReport_0']");
                                    Thread.Sleep(2000);

                                    //FORMA REPORTE DETALLADO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblNivEnde_1']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Nomina Detallado", true, file);

                                    //GENERAR
                                    selenium.Click("//a[@id='ctl00_btnGenerar']");
                                    Thread.Sleep(8000);
                                    Screenshot("Reporte Endeudamiento Generado", true, file);
                                    Thread.Sleep(6000);
                                    //Abrir pdf
                                    if (database == "ORA")
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220510163355397.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);
                                    }
                                    else
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220509171500506.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);

                                    }

                                    Screenshot("PDF ABIERTO", true, file);
                                    Thread.Sleep(20000);
                                    LimpiarProcesos();
                                    Thread.Sleep(2000);
                                    selenium.Close();

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
        public void ReporteNivelEndeudamientoWebConfigNominaResumido()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.ReporteNivelEndeudamientoWebConfigNominaResumido")
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

                            if (
                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();

                                try
                                {

                                    string database = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";

                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";

                                    }

                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";

                                    }
                                    if (url.ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";


                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Scroll("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Reportes", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(1500);
                                    //CHECK NIVEL ENDEUDAMIENTO
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(2000);
                                    //fechas
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlFecDesdA']", "2021");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    //TIPO DE REPORTE NOMINA
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblTipReport_0']");
                                    Thread.Sleep(2000);
                                    //FORMA REPORTE RESUMIDO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblNivEnde_0']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Nomina Resumido", true, file);

                                    //GENERAR
                                    selenium.Click("//a[@id='ctl00_btnGenerar']");
                                    Thread.Sleep(8000);
                                    Screenshot("Reporte Endeudamiento Generado", true, file);
                                    Thread.Sleep(6000);
                                    //Abrir pdf
                                    if (database == "ORA")
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220511164830312.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);
                                    }
                                    else
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220509171511816.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);

                                    }

                                    Screenshot("PDF ABIERTO", true, file);
                                    Thread.Sleep(20000);
                                    LimpiarProcesos();
                                    Thread.Sleep(2000);
                                    selenium.Close();

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
        public void ReporteNivelEndeudamientoWebConfigPrimaDetallado()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.ReporteNivelEndeudamientoWebConfigPrimaDetallado")
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

                            if (
                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();

                                try
                                {

                                    string database = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";

                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";

                                    }

                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";

                                    }
                                    if (url.ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";


                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Scroll("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Reportes", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(1500);
                                    //CHECK NIVEL ENDEUDAMIENTO
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(2000);
                                    //fechas
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlFecDesdA']", "2021");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    //TIPO DE REPORTE PRIMA
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblTipReport_1']");
                                    Thread.Sleep(2000);
                                    //FORMA REPORTE DETALLADO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblNivEnde_1']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Prima Detallado", true, file);
                                    Thread.Sleep(2000);

                                    //GENERAR
                                    selenium.Click("//a[@id='ctl00_btnGenerar']");
                                    Thread.Sleep(8000);
                                    Screenshot("Reporte Endeudamiento Generado", true, file);
                                    Thread.Sleep(6000);

                                    //Abrir pdf
                                    if (database == "ORA")
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220511170848710.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);
                                    }
                                    else
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220509171521481.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);

                                    }

                                    Screenshot("PDF ABIERTO", true, file);
                                    Thread.Sleep(20000);
                                    LimpiarProcesos();
                                    Thread.Sleep(2000);
                                    selenium.Close();

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
        public void ReporteNivelEndeudamientoWebConfigPrimaResumido()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.ReporteNivelEndeudamientoWebConfigPrimaResumido")
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

                            if (
                                //Datos Login
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();

                                try
                                {

                                    string database = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";

                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";

                                    }

                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";

                                    }
                                    if (url.ToLower() == "http://ophtsph:8085/".ToLower())
                                    {
                                        database = "SQL";


                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Scroll("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Mis Reportes')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Reportes", true, file);

                                    Thread.Sleep(1000);

                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(1500);
                                    //CHECK NIVEL ENDEUDAMIENTO
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(2000);
                                    //fechas
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlFecDesdA']", "2021");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas", true, file);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_chkNivEnde')]");
                                    Thread.Sleep(1000);
                                    //TIPO DE REPORTE PRIMA
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblTipReport_1']");
                                    Thread.Sleep(2000);
                                    //FORMA REPORTE RESUMIDO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblNivEnde_0']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Prima Resumido", true, file);
                                    Thread.Sleep(2000);

                                    //GENERAR
                                    selenium.Click("//a[@id='ctl00_btnGenerar']");
                                    Thread.Sleep(8000);
                                    Screenshot("Reporte Endeudamiento Generado", true, file);
                                    Thread.Sleep(6000);
                                    //Abrir pdf
                                    if (database == "ORA")
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220511165601184.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);
                                    }
                                    else
                                    {
                                        string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/kactusscm/Downloads/ArchivoKNmRnien_20220509171534225.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(60000);

                                    }

                                    Screenshot("PDF ABIERTO", true, file);
                                    Thread.Sleep(20000);
                                    LimpiarProcesos();
                                    Thread.Sleep(2000);
                                    selenium.Close();

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
        public void NovedadesTemporalesRegistroNuevoRRHH()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_Cafam.NovedadesTemporalesRegistroNuevoRRHH")
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

                            if (
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["JefeUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["JefePass"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                rows["Cantidad"].ToString().Length != 0 && rows["Cantidad"].ToString() != null &&
                                rows["ValCuota"].ToString().Length != 0 && rows["ValCuota"].ToString() != null &&
                                rows["NumCuotas"].ToString().Length != 0 && rows["NumCuotas"].ToString() != null &&
                                rows["ValTotal"].ToString().Length != 0 && rows["ValTotal"].ToString() != null &&
                                rows["CodConc"].ToString().Length != 0 && rows["CodConc"].ToString() != null &&
                                rows["TipApli"].ToString().Length != 0 && rows["TipApli"].ToString() != null &&
                                rows["IndActi"].ToString().Length != 0 && rows["IndActi"].ToString() != null &&
                                rows["IndActi2"].ToString().Length != 0 && rows["IndActi2"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Concepto = rows["Concepto"].ToString();
                                string Cantidad = rows["Cantidad"].ToString();
                                string ValCuota = rows["ValCuota"].ToString();
                                string NumCuotas = rows["NumCuotas"].ToString();
                                string ValTotal = rows["ValTotal"].ToString();
                                string CodConc = rows["CodConc"].ToString();
                                string TipApli = rows["TipApli"].ToString();
                                string IndActi = rows["IndActi"].ToString();
                                string IndActi2 = rows["IndActi2"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();

                                try
                                {
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //PARAMETRIZACION
                                    if (database == "SQL")
                                    {
                                        string eliminarRegistro = $"DELETE NM_NOVTE WHERE COD_EMPR = '{CodEmpr}' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro2 = $"DELETE NM_NOVED WHERE COD_EMPR = '{CodEmpr}' AND COD_EMPL = '{EmpleadoUser}' AND COD_CONC = '{CodConc}'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro3 = $"DELETE NM_SOLTR where COD_RESP='{EmpleadoUser}' AND TIP_APLI='T'";
                                        db.UpdateDeleteInsert(eliminarRegistro3, database, user);
                                    }
                                    else
                                    {
                                        string eliminarRegistro = $"DELETE NM_NOVTE WHERE COD_EMPR = '421' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro2 = $"DELETE NM_NOVED WHERE COD_EMPR = '421' AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro3 = $"DELETE NM_SOLTR where COD_RESP='{EmpleadoUser}' AND TIP_APLI='T'";
                                        db.UpdateDeleteInsert(eliminarRegistro3, database, user);
                                    }
                                    //INICIO PRUEBA                                 
                                    //RRHH
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "ORA")
                                    {
                                        selenium.Click("//button[contains(.,'GESTION HUMANA')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//button[contains(.,'Rol RRHH')]");
                                    }
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol RRHH", true, file);

                                    selenium.Scroll("//a[contains(.,'PUESTO DE TRABAJO')]");
                                    selenium.Click("//a[contains(.,'PUESTO DE TRABAJO')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Novedades Temporales RRHH')]");
                                    selenium.Click("//a[contains(.,'Novedades Temporales RRHH')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Novedades Temporales RRHH", true, file);

                                    //NUEVO
                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nuevo", true, file);
                                    //BUSCAR EMPLEADO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtConsuCedEmpl']");
                                    Thread.Sleep(1000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtConsuCedEmpl']", EmpleadoUser);
                                    selenium.Screenshot("Empleado a Buscar", true, file);
                                    //CLICK DETALLE
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsulCed']");
                                    Thread.Sleep(5000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_dgBiEmpl_ctl03_btnVer']");
                                    Thread.Sleep(5000);
                                    //concepto
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomConc']", Concepto);
                                    Thread.Sleep(2000);
                                    //cantidad
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtCanNove']", Cantidad);
                                    Thread.Sleep(2000);
                                    //Cuotas
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValCuot']", ValCuota);
                                    Thread.Sleep(2000);
                                    //numero cuotas
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNumCuot']",NumCuotas);
                                    Thread.Sleep(2000);
                                    //Valor total
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtValTota']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValTota']", ValTotal);
                                    Thread.Sleep(2000);
                                    //Saldo
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtSalNove']", ValTotal);
                                    Thread.Sleep(2000);
                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Novedad Registrada", true, file);
                                    Thread.Sleep(6000);
                                    selenium.Close();


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

