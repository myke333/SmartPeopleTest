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
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Chrome;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Management;
using OpenQA.Selenium.Interactions;
using APITest;


namespace Web_Kactus_Test
{
    /// <summary>
    /// Descripción resumida de CodedUITest1
    /// </summary>
    [CodedUITest]
    public class Reclutamiento : FuncionesVitales
    {
        string app = "Reclutamiento";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();
        public Reclutamiento()
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
        public void DatosBasicos() 
        {
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
           // TableOrder = "ktes1";

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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Reclutamiento.DatosBasicos")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RLPassword"].ToString().Length != 0 && rows["RLPassword"].ToString() != null &&
                                // Data Trayectoria RL/////////////////////
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["Perfil"].ToString().Length != 0 && rows["Perfil"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null 
                                
                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RLPassword"].ToString();
                                // Data Trayectoria RL/////////////////////
                                string Nombre = rows["Nombre"].ToString();
                                string Perfil = rows["Perfil"].ToString();
                                string Url = rows["URL"].ToString();


                                try
                                {
                                    string database = "";
                                    if (Url.ToLower().Contains("ora"))
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
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();
                                    //INICIO
                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    //DATOS BASICOS
                                    selenium.Click("//*[@id='MenuContex']/div[2]/div[1]/ul/li[2]/a");                              
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Básicos", true, file);
                                    //NOMBRE
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_txtNomEmpl']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomEmpl']", Nombre);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro Nombre", true, file);
                                    //PERFIL LABORAL
                                    selenium.ScrollTo("0", "800");
                                    Thread.Sleep(2000);
                                    selenium.Clear("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValMPerfile_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValMPerfile_txtTexto']", Perfil);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro Perfil Laboral", true, file);
                                    //AREA EXPERIENCIA
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAreInte']", "ADMINISTRATIVA Y FINANCIERA");
                                    Thread.Sleep(2000);
                                    selenium.ScrollTo("0", "900");
                                    selenium.Screenshot("Registro Area Experiencia", true, file);
                                    //ACTUALIZAR
                                    selenium.Click("//a[@id='btnActualizar']");
                                    Thread.Sleep(2000);
                                    if (selenium.ExistControl("//span[contains(@id,'lblError')]"))
                                    {
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Actualización exitosa", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                        Thread.Sleep(2000);
                                        selenium.Scroll("//a[contains(text(),'Eliminar')]");
                                        Thread.Sleep(2000);
                                        selenium.Click("//a[contains(text(),'Eliminar')]");

                                    }else

                                    {
                                        Thread.Sleep(500);
                                        selenium.Screenshot("No se pudo actualizar correctamente", true, file);

                                        errorMessagesMetodo.Add("El mensaje de confimación de actualización no es correcto, revise por favor");
                                    }

                                    //////
                                    ConverWordToPDF(file, database);
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
        public void Documentos()
        {
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
            //TableOrder = "KTES1";

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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Reclutamiento.Documentos")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null &&
                                // Data Documentos  RL/////////////////////
                                rows["DddlNomDocus"].ToString().Length != 0 && rows["DddlNomDocus"].ToString() != null &&
                                rows["DtxtNumDocu"].ToString().Length != 0 && rows["DtxtNumDocu"].ToString() != null &&
                                rows["DddlPais"].ToString().Length != 0 && rows["DddlPais"].ToString() != null &&
                                rows["DddlDeps"].ToString().Length != 0 && rows["DddlDeps"].ToString() != null &&
                                rows["DddlMuns"].ToString().Length != 0 && rows["DddlMuns"].ToString() != null &&
                                rows["DtxtFecha"].ToString().Length != 0 && rows["DtxtFecha"].ToString() != null &&
                                rows["Observaciones"].ToString().Length != 0 && rows["Observaciones"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                // Data Documentos  RL/////////////////////
                                string DddlNomDocus = rows["DddlNomDocus"].ToString();
                                string DtxtNumDocu = rows["DtxtNumDocu"].ToString();
                                string DddlPais = rows["DddlPais"].ToString();
                                string DddlDeps = rows["DddlDeps"].ToString();
                                string DddlMuns = rows["DddlMuns"].ToString();
                                string DtxtFecha = rows["DtxtFecha"].ToString();
                                string Observaciones = rows["Observaciones"].ToString();
                                string Url = rows["Url"].ToString();
                                string Maquina = rows["Maquina"].ToString();

                                try
                                {
                                    string database = "";
                                    if (Url.ToLower().Contains("ora"))
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
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();

                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);

                                    //Process: Documentos ///////////////////////////////////////////////////////////////////
                                    selenium.Click("//a[contains(text(),'Documentos')]");
                                    //EDICIÓN
                                    selenium.Click("//a[contains(text(),'Detalle')]");
                                    selenium.Screenshot("Edición Documentos", true, file);
                                    selenium.Clear("//input[contains(@id,'ContenidoPagina_txtObsErva')]");
                                    selenium.SendKeys("//input[contains(@id,'ContenidoPagina_txtObsErva')]", Observaciones);
;                                   selenium.Click("//button[contains(@id,'btnActualizar')]");
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Actualización", true, file);
                                    
                                    //NUEVO
                                    selenium.Click("//button[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Documentos", true, file);

                                    Thread.Sleep(700);
                                    selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlNomDocu')]", DddlNomDocus);
                                    Thread.Sleep(700);

                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtNumDocu')]", DtxtNumDocu);
                                    Thread.Sleep(700);

                                    selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_KCtrDivipPais_ddlPai')]", DddlPais);
                                    Thread.Sleep(700);
                                    selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_KCtrDivipPais_ddlDep')]", DddlDeps);
                                    Thread.Sleep(700);
                                    selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_KCtrDivipPais_ddlMun')]", DddlMuns);
                                    Thread.Sleep(700);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlFecExpe_txtFecha')]", DtxtFecha);
                                    Thread.Sleep(700);

                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecVenc_txtFecha']", "17/03/2018");
                                    Thread.Sleep(700);

                                    selenium.Screenshot("Datos Documentos", true, file);


                                    selenium.Click("//button[contains(@id,'btnGuardar')]");

                                    if(selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_grvRlEmpdo_ctl02_LinkButton2']"))                                    
                                    {
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Existe Datos Documentos", true, file);

                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_grvRlEmpdo_ctl02_LinkButton2']");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Elimina Datos Documentos", true, file);

                                    }
                                    else
                                    {
                                        Thread.Sleep(500);
                                        selenium.Screenshot("No existe Datos Documentos", true, file);

                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se inserto el documento");
                                    }
                                    //////
                                    ConverWordToPDF(file, database);
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
        public void EducacionFormal()
        {
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Reclutamiento.EducacionFormal")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null &&
                                // Data Educacion Formal  RL///////////
                                rows["RlModalidad"].ToString().Length != 0 && rows["RlModalidad"].ToString() != null &&
                                rows["RlNomEstudios"].ToString().Length != 0 && rows["RlNomEstudios"].ToString() != null &&
                                rows["txtNomEspp"].ToString().Length != 0 && rows["txtNomEspp"].ToString() != null &&
                                rows["txtNomInst"].ToString().Length != 0 && rows["txtNomInst"].ToString() != null &&
                                rows["txtTieEstu"].ToString().Length != 0 && rows["txtTieEstu"].ToString() != null &&
                                rows["ddlUniTiems"].ToString().Length != 0 && rows["ddlUniTiems"].ToString() != null &&
                                rows["KCtrlFecInic"].ToString().Length != 0 && rows["KCtrlFecInic"].ToString() != null &&
                                rows["KCtrlFecTerm"].ToString().Length != 0 && rows["KCtrlFecTerm"].ToString() != null &&
                                rows["ddlPais"].ToString().Length != 0 && rows["ddlPais"].ToString() != null &&
                                rows["ddlDeps"].ToString().Length != 0 && rows["ddlDeps"].ToString() != null &&
                                rows["PromedioEdita"].ToString().Length != 0 && rows["PromedioEdita"].ToString() != null &&
                                rows["ddlMuns"].ToString().Length != 0 && rows["ddlMuns"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                // Data Educacion Formal  RL///////////                                             
                                string RlModalidad = rows["RlModalidad"].ToString();
                                string RlNomEstudios = rows["RlNomEstudios"].ToString();
                                string txtNomEspp = rows["txtNomEspp"].ToString();
                                string txtNomInst = rows["txtNomInst"].ToString();
                                string txtTieEstu = rows["txtTieEstu"].ToString();
                                string ddlUniTiems = rows["ddlUniTiems"].ToString();
                                string KCtrlFecInic = rows["KCtrlFecInic"].ToString();
                                string ddlPais = rows["ddlPais"].ToString();
                                string ddlDeps = rows["ddlDeps"].ToString();
                                string ddlMuns = rows["ddlMuns"].ToString();
                                string KCtrlFecTerm = rows["KCtrlFecTerm"].ToString();
                                string PromedioEdita = rows["PromedioEdita"].ToString();
                                string Url = rows["Url"].ToString();
                                string Maquina = rows["Maquina"].ToString();

                                try
                                {
                                    string database = "";
                                    if (Url.ToLower().Contains("ora"))
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
                                    DateTime dateAndTime = DateTime.Now;
                                    string datetime = dateAndTime.ToString("ddMMyyyy_HHmmss");
                                    string UpdateData = "TEST_" + datetime;
                                    // Data Basicos RL//////
                                    string NomAspirante = UpdateData;
                                    ///////////////////////////////////////

                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();

                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);

                                    //Process: Educacion Fromal//////////////////////////////////////////////////////////
                                    selenium.Click("//a[contains(text(),'Edu. Formal')]");
                                    selenium.Screenshot("Educación Formal", true, file);


                                    //EDITAR REGISTRO
                                    selenium.Click("//a[contains(text(),'Detalle')]");
                                    selenium.Screenshot("Datos Educación Formal", true, file);

                                    selenium.Clear("//input[contains(@id,'txtProCarr')]");

                                    selenium.SendKeys("//input[contains(@id,'txtProCarr')]", PromedioEdita);
                                    selenium.Screenshot("Edición Educación Formal", true, file);

                                    selenium.Click("//button[contains(@id,'btnActualizar')]");

                                    //NUEVO
                                    selenium.Click("//button[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Educación Formal", true, file);

                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomModi')]",RlModalidad);
                                    Thread.Sleep(2000);

                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomEstu')]",RlNomEstudios);
                                    selenium.SendKeys("//input[contains(@id,'txtNomEspp')]",txtNomEspp);

                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomInst']", txtNomInst);


                                    selenium.SendKeys("//input[contains(@id,'txtTieEstu')]",txtTieEstu);

                                    selenium.SelectElementByName("//select[contains(@id,'ddlUniTiem')]",ddlUniTiems);

                                    selenium.SendKeys("//input[contains(@id,'KCtrlFecInic_txtFecha')]",KCtrlFecInic);
                                    selenium.SendKeys("//input[contains(@id,'KCtrlFecTerm_txtFecha')]",KCtrlFecTerm);

                                    selenium.SelectElementByName("//select[contains(@id,'ddlPai')]",ddlPais);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlDep')]",ddlDeps);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlMun')]",ddlMuns);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Datos Educación Formal", true, file);

                                    selenium.Click("//button[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(500);
                                    try
                                    {
                                        selenium.AcceptAlert();
                                    }
                                    catch { }

                                    if(selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_grvRlEdfor_ctl02_LinkButton2']"))
                                    {
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Existe Educación Formal", true, file);

                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_grvRlEdfor_ctl02_LinkButton2']");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Elimina Educación Formal", true, file);

                                    }
                                    else
                                    {
                                        Thread.Sleep(500);
                                        selenium.Screenshot("No existe Educación Formal", true, file);

                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se inserto el estudio formal");
                                    }
                                    //////
                                    ConverWordToPDF(file, database);
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
        public void EducacionNoFormal()
        {
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Reclutamiento.EducacionNoFormal")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null &&
                                // Data Educacion no Formal  RL////////
                                rows["ddlNomModis"].ToString().Length != 0 && rows["ddlNomModis"].ToString() != null &&
                                rows["ddlNomEstus"].ToString().Length != 0 && rows["ddlNomEstus"].ToString() != null &&
                                rows["txtNomEspe"].ToString().Length != 0 && rows["txtNomEspe"].ToString() != null &&
                                rows["txtNomInsts"].ToString().Length != 0 && rows["txtNomInsts"].ToString() != null &&
                                rows["txtFechaIni"].ToString().Length != 0 && rows["txtFechaIni"].ToString() != null &&
                                rows["txtFechaFin"].ToString().Length != 0 && rows["txtFechaFin"].ToString() != null &&
                                rows["txtTieEstuTime"].ToString().Length != 0 && rows["txtTieEstuTime"].ToString() != null &&
                                rows["ddlUniTiemDis"].ToString().Length != 0 && rows["ddlUniTiemDis"].ToString() != null &&
                                rows["KCtrDivipPaisddlPais"].ToString().Length != 0 && rows["KCtrDivipPaisddlPais"].ToString() != null &&
                                rows["KCtrDivipPaisddlDep"].ToString().Length != 0 && rows["KCtrDivipPaisddlDep"].ToString() != null &&
                                rows["KCtrDivipPaisddlMun"].ToString().Length != 0 && rows["KCtrDivipPaisddlMun"].ToString() != null &&
                                rows["txtTieEstuTimeEdita"].ToString().Length != 0 && rows["txtTieEstuTimeEdita"].ToString() != null &&
                                rows["ddlUniTiemDisEdit"].ToString().Length != 0 && rows["ddlUniTiemDisEdit"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                // Data Educacion no Formal  RL////////
                                string ddlNomModis = rows["ddlNomModis"].ToString();
                                string ddlNomEstus = rows["ddlNomEstus"].ToString();
                                string txtNomEspe = rows["txtNomEspe"].ToString();
                                string txtNomInsts = rows["txtNomInsts"].ToString();
                                string txtFechaIni = rows["txtFechaIni"].ToString();
                                string txtFechaFin = rows["txtFechaFin"].ToString();
                                string txtTieEstuTime = rows["txtTieEstuTime"].ToString();
                                string ddlUniTiemDis = rows["ddlUniTiemDis"].ToString();
                                string KCtrDivipPaisddlPais = rows["KCtrDivipPaisddlPais"].ToString();
                                string KCtrDivipPaisddlDep = rows["KCtrDivipPaisddlDep"].ToString();
                                string KCtrDivipPaisddlMun = rows["KCtrDivipPaisddlMun"].ToString();
                                string txtTieEstuTimeEdita = rows["txtTieEstuTimeEdita"].ToString();
                                string ddlUniTiemDisEdit = rows["ddlUniTiemDisEdit"].ToString();
                                string Url = rows["Url"].ToString();
                                string Maquina = rows["Maquina"].ToString();

                                try
                                {
                                    string database = "";
                                    if (Url.ToLower().Contains("ora"))
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
                                    DateTime dateAndTime = DateTime.Now;
                                    string datetime = dateAndTime.ToString("ddMMyyyy_HHmmss");
                                    string UpdateData = "TEST_" + datetime;
                                    // Data Basicos RL//////
                                    string NomAspirante = UpdateData;
                                    ///////////////////////////////////////

                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();

                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);
                                    
                                    //Process: Educacion no Fromal////////////////////////////////////////////////////////
                                    selenium.Click("//a[contains(text(),'Edu. No Formal')]");
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Edu. No Formal", true, file);

                                    //EDICIÓN
                                    selenium.Click("//a[contains(text(),'Detalle')]");
                                    selenium.Clear("//input[contains(@id,'txtTieEstu')]");
                                    selenium.SendKeys("//input[contains(@id,'txtTieEstu')]",txtTieEstuTimeEdita);                                  
                                    selenium.SelectElementByName("//select[contains(@id,'ddlUniTiem')]", ddlUniTiemDisEdit);
                                    selenium.Screenshot("Actualización de Datos", true, file);
                                    selenium.Click("//button[contains(@id,'btnActualizar')]");

                                    //NUEVO
                                    selenium.Click("//button[contains(@id,'btnNuevo')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomModi')]", ddlNomModis);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomEstu')]", ddlNomEstus);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtNomEspe')]",txtNomEspe);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtNomInst')]",txtNomInsts);
                                    Thread.Sleep(2000);

                                    selenium.SendKeys("//input[contains(@id,'KCtrlFecInic_txtFecha')]",txtFechaIni);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'KCtrlFecTerm_txtFecha')]",txtFechaFin);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtTieEstu')]",txtTieEstuTime);
                                    Thread.Sleep(1500);


                                    selenium.SelectElementByName("//select[contains(@id,'ddlUniTiem')]",ddlUniTiemDis);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlPai')]", KCtrDivipPaisddlPais);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlDep')]", KCtrDivipPaisddlDep);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlMun')]", KCtrDivipPaisddlMun);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Datos Edu. No Formal", true, file);
                                    Thread.Sleep(1500);
                                    selenium.Click("//button[contains(@id,'btnGuardar')]");
                                    try
                                    {
                                        Thread.Sleep(500);
                                        selenium.AcceptAlert();
                                    }
                                    catch { }
                                    selenium.Screenshot("Existe Edu. No Formal", true, file);

                                    //ELIMINACIÓN
                                    if(selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_grvRlEdnfo_ctl02_LinkButton2']"))
                                    {
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Existe Educación No Formal", true, file);

                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_grvRlEdnfo_ctl02_LinkButton2']");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Elimina Educación No Formal", true, file);

                                    }
                                    else
                                    {
                                        Thread.Sleep(500);
                                        selenium.Screenshot("No existe Educación No Formal", true, file);

                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se inserto el estudio formal");
                                    }
                                    //////
                                    ConverWordToPDF(file, database);
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
        public void ExperienciaLaboral()
        {
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Reclutamiento.ExperienciaLaboral")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null &&
                                // Data Experiencia Laboral //
                                rows["ElNomEmpr"].ToString().Length != 0 && rows["ElNomEmpr"].ToString() != null &&
                                rows["EltxtDirEmpr"].ToString().Length != 0 && rows["EltxtDirEmpr"].ToString() != null &&
                                rows["EltxtTelEmpr"].ToString().Length != 0 && rows["EltxtTelEmpr"].ToString() != null &&
                                rows["ElddlTipEmprs"].ToString().Length != 0 && rows["ElddlTipEmprs"].ToString() != null &&
                                rows["EltxtFecha"].ToString().Length != 0 && rows["EltxtFecha"].ToString() != null &&
                                rows["EltxtFechas"].ToString().Length != 0 && rows["EltxtFechas"].ToString() != null &&
                                rows["EltxtCarDese"].ToString().Length != 0 && rows["EltxtCarDese"].ToString() != null &&
                                rows["ElddlDedIcacs"].ToString().Length != 0 && rows["ElddlDedIcacs"].ToString() != null &&
                                rows["ElddlTipConts"].ToString().Length != 0 && rows["ElddlTipConts"].ToString() != null &&
                                rows["ElddlMotRetis"].ToString().Length != 0 && rows["ElddlMotRetis"].ToString() != null &&
                                rows["ElddlManPerss"].ToString().Length != 0 && rows["ElddlManPerss"].ToString() != null &&
                                rows["ElddlPais"].ToString().Length != 0 && rows["ElddlPais"].ToString() != null &&
                                rows["ElddlDeps"].ToString().Length != 0 && rows["ElddlDeps"].ToString() != null &&
                                rows["ElddlMuns"].ToString().Length != 0 && rows["ElddlMuns"].ToString() != null &&
                                rows["EltxtJefInme"].ToString().Length != 0 && rows["EltxtJefInme"].ToString() != null &&
                                rows["EltxtCarJefe"].ToString().Length != 0 && rows["EltxtCarJefe"].ToString() != null &&
                                rows["ElddlAreExpes"].ToString().Length != 0 && rows["ElddlAreExpes"].ToString() != null &&
                                rows["EltxtEntMail"].ToString().Length != 0 && rows["EltxtEntMail"].ToString() != null &&
                                rows["ddlCodActis"].ToString().Length != 0 && rows["ddlCodActis"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                string Url = rows["Url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                // Data Experiencia Laboral //
                                string ElNomEmpr = rows["ElNomEmpr"].ToString();
                                string EltxtDirEmpr = rows["EltxtDirEmpr"].ToString();
                                string EltxtTelEmpr = rows["EltxtTelEmpr"].ToString();
                                string ElddlTipEmprs = rows["ElddlTipEmprs"].ToString();
                                string EltxtFecha = rows["EltxtFecha"].ToString();
                                string EltxtFechas = rows["EltxtFechas"].ToString();
                                string EltxtCarDese = rows["EltxtCarDese"].ToString();
                                string ElddlDedIcacs = rows["ElddlDedIcacs"].ToString();
                                string ElddlTipConts = rows["ElddlTipConts"].ToString();
                                string ElddlMotRetis = rows["ElddlMotRetis"].ToString();
                                string ElddlManPerss = rows["ElddlManPerss"].ToString();
                                string ElddlPais = rows["ElddlPais"].ToString();
                                string ElddlDeps = rows["ElddlDeps"].ToString();
                                string ElddlMuns = rows["ElddlMuns"].ToString();
                                string EltxtJefInme = rows["EltxtJefInme"].ToString();
                                string EltxtCarJefe = rows["EltxtCarJefe"].ToString();
                                string ElddlAreExpes = rows["ElddlAreExpes"].ToString();
                                string EltxtEntMail = rows["EltxtEntMail"].ToString();
                                string ddlCodActis = rows["ddlCodActis"].ToString();

                                try
                                {
                                    string database = "";
                                    if (Url.ToLower().Contains("ora"))
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
                                    DateTime dateAndTime = DateTime.Now;
                                    string datetime = dateAndTime.ToString("ddMMyyyy_HHmmss");
                                    string UpdateData = "TEST_" + datetime;
                                    // Data Basicos RL//////
                                    string NomAspirante = UpdateData;
                                    ///////////////////////////////////////

                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();



                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);

                                    //Process: 	Experiencia Laboral //////////////////////////////////////////////////////
                                    selenium.Click("//a[contains(text(),'Exp Laboral')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Experiencia laboral", true, file);

                                    Thread.Sleep(1500);

                                    selenium.Click("//button[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtNomEmpr')]", ElNomEmpr);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtDirEmpr')]", EltxtDirEmpr);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipEmpr')]", ElddlTipEmprs);
                                    Thread.Sleep(1500);

                                    selenium.SendKeys("//input[contains(@id,'txtEntMail')]", EltxtEntMail);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'KCtrlFecIngr_txtFecha')]", EltxtFecha);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'KCtrlFecReti_txtFecha')]", EltxtFechas);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtCarDese')]", EltxtCarDese);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtTelEmpr')]", EltxtTelEmpr);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlDedIcac')]", ElddlDedIcacs);
                                    Thread.Sleep(1500);

                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ElddlTipConts);
                                    Thread.Sleep(1500);

                                    selenium.SelectElementByName("//select[contains(@id,'ddlMotReti')]", ElddlMotRetis);
                                    Thread.Sleep(1500);

                                    selenium.SelectElementByName("//select[contains(@id,'ddlManPers')]", ElddlManPerss);
                                    Thread.Sleep(1500);

                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodActi')]", ddlCodActis);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlPai')]", ElddlPais);
                                    Thread.Sleep(1500);

                                    selenium.SelectElementByName("//select[contains(@id,'ddlDep')]", ElddlDeps);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlMun')]", ElddlMuns);
                                    Thread.Sleep(1500);

                                    selenium.SendKeys("//input[contains(@id,'txtJefInme')]", EltxtJefInme);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtCarJefe')]", EltxtCarJefe);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlAreExpe')]", ElddlAreExpes);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Datos Experiencia Laboral", true, file);
                                    Thread.Sleep(1500);
                                    selenium.Click("//button[contains(@id,'btnGuardar')]");

                                
                                    //ELIMINACIÓN


                                if (selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_grvRlHvext_ctl02_LinkButton2']"))
                                {
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Existe Exp Laboral", true, file);

                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_grvRlHvext_ctl02_LinkButton2']");
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Elimina Exp Laboral", true, file);

                                }
                                else
                                {
                                    Thread.Sleep(500);
                                    selenium.Screenshot("No existe Exp Laboral", true, file);

                                    errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se inserto la experiencia laboral");
                                }
                            
                                        ////////////////////////////////////////////////////////////////////////////////////
                                    
                                    //////
                                    ConverWordToPDF(file, database);
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
        public void Familiares()
        {
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Reclutamiento.Familiares")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null &&
                                // Data Familiares  RL/////////////////
                                rows["FatxtCodFami"].ToString().Length != 0 && rows["FatxtCodFami"].ToString() != null &&
                                rows["FaddlTipIdens"].ToString().Length != 0 && rows["FaddlTipIdens"].ToString() != null &&
                                rows["FatxtNomFami1"].ToString().Length != 0 && rows["FatxtNomFami1"].ToString() != null &&
                                rows["FaddlTipRelas"].ToString().Length != 0 && rows["FaddlTipRelas"].ToString() != null &&
                                rows["FatxtApeFami1"].ToString().Length != 0 && rows["FatxtApeFami1"].ToString() != null &&
                                rows["FaddlSexFamis"].ToString().Length != 0 && rows["FaddlSexFamis"].ToString() != null &&
                                rows["FaddlFecNaciAs"].ToString().Length != 0 && rows["FaddlFecNaciAs"].ToString() != null &&
                                rows["FaddlFecNaciMs"].ToString().Length != 0 && rows["FaddlFecNaciMs"].ToString() != null &&
                                rows["FaddlFecNaciDs"].ToString().Length != 0 && rows["FaddlFecNaciDs"].ToString() != null &&
                                rows["FaddlGruSangs"].ToString().Length != 0 && rows["FaddlGruSangs"].ToString() != null &&
                                rows["FaddlFacSangs"].ToString().Length != 0 && rows["FaddlFacSangs"].ToString() != null &&
                                rows["FaddlEstCivis"].ToString().Length != 0 && rows["FaddlEstCivis"].ToString() != null &&
                                rows["Hobbies"].ToString().Length != 0 && rows["Hobbies"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                // Data Familiares  RL/////////////////
                                string FatxtCodFami = rows["FatxtCodFami"].ToString();
                                string FaddlTipIdens = rows["FaddlTipIdens"].ToString();
                                string FatxtNomFami1 = rows["FatxtNomFami1"].ToString();
                                string FaddlTipRelas = rows["FaddlTipRelas"].ToString();
                                string FatxtApeFami1 = rows["FatxtApeFami1"].ToString();
                                string FaddlSexFamis = rows["FaddlSexFamis"].ToString();
                                string FaddlFecNaciAs = rows["FaddlFecNaciAs"].ToString();
                                string FaddlFecNaciMs = rows["FaddlFecNaciMs"].ToString();
                                string FaddlFecNaciDs = rows["FaddlFecNaciDs"].ToString();
                                string FaddlGruSangs = rows["FaddlGruSangs"].ToString();
                                string FaddlFacSangs = rows["FaddlFacSangs"].ToString();
                                string FaddlEstCivis = rows["FaddlEstCivis"].ToString();
                                string Hobbies = rows["Hobbies"].ToString();
                                string Url = rows["Url"].ToString();
                                string Maquina = rows["Maquina"].ToString();

                                try
                                {
                                    string database = "";
                                    if (Url.ToLower().Contains("ora"))
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
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();

                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);


                                    //Process: 	Familiares //////////////////////////////////////////////////////
                                    selenium.Click("//a[contains(text(),'Familiares')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Familiares", true, file);

                                    //EDICION
                                    selenium.Click("//input[contains(@id,'ctl03_LinkButton1')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtHobFami')]",Hobbies);
                                    selenium.Screenshot("Actualización de Familiares", true, file);
                                    Thread.Sleep(1500);
                                    selenium.Click("//button[contains(@id,'btnActualizar')]");
                                    
                                    //NUEVO
                                    selenium.Click("//button[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtCodFami')]",FatxtCodFami);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipIden')]",FaddlTipIdens);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtNomFami1')]",FatxtNomFami1);
                                    Thread.Sleep(1500);

                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipRela')]",FaddlTipRelas);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtApeFami1')]",FatxtApeFami1);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlSexFami')]",FaddlSexFamis);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlFecNaciA')]",FaddlFecNaciAs);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlFecNaciM')]",FaddlFecNaciMs);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlFecNaciD')]",FaddlFecNaciDs);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlGruSang')]",FaddlGruSangs);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlFacSang')]",FaddlFacSangs);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlEstCivi')]",FaddlEstCivis);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Datos familiares", true, file);

                                    selenium.Click("//button[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Registro de Datos Familiares", true, file);

                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_dtgEdFamil_ctl03_LinkButton1')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//button[contains(@id,'btnEliminar')]");
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Eliminación Datos Familiares", true, file);

                                    //////
                                    ConverWordToPDF(file, database);
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
        public void Idiomas()
        {
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Reclutamiento.Idiomas")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null &&
                                // Data Idiomas  RL///////////////////
                                rows["IddlNomIdios"].ToString().Length != 0 && rows["IddlNomIdios"].ToString() != null &&
                                rows["IddlHabIdios"].ToString().Length != 0 && rows["IddlHabIdios"].ToString() != null &&
                                rows["IddlLeeIdios"].ToString().Length != 0 && rows["IddlLeeIdios"].ToString() != null &&
                                rows["IddlEscIdios"].ToString().Length != 0 && rows["IddlEscIdios"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                // Data Idiomas  RL///////////////////
                                string IddlNomIdios = rows["IddlNomIdios"].ToString();
                                string IddlHabIdios = rows["IddlHabIdios"].ToString();
                                string IddlLeeIdios = rows["IddlLeeIdios"].ToString();
                                string IddlEscIdios = rows["IddlEscIdios"].ToString();
                                string Url = rows["Url"].ToString();
                                string Maquina = rows["Maquina"].ToString();

                                try
                                {
                                    string database = "";
                                    if (Url.ToLower().Contains("ora"))
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
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();

                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);

                                    //Process: Idiomas ///////////////////////////////////////////////////////////////////
                                    selenium.Click("//a[contains(text(),'Idiomas')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Idiomas", true, file);

                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomIdio')]",IddlNomIdios);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlHabIdio')]",IddlHabIdios);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlLeeIdio')]",IddlLeeIdios);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlEscIdio')]",IddlEscIdios);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Datos idiomas", true, file);

                                    selenium.Click("//button[contains(@id,'btnGuardar')]");
                                    if(selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_grvRlEmidi_ctl02_LinkButton1']"))
                                    {
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Existe Datos idiomas", true, file);

                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_grvRlEmidi_ctl02_LinkButton1']");
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("elimina Datos idiomas", true, file);

                                    }
                                    else
                                    {
                                        Thread.Sleep(500);
                                        selenium.Screenshot("no existe Datos idiomas", true, file);

                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se inserto el idioma");
                                    }
                                    //////
                                    ConverWordToPDF(file, database);
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
        public void MisTallas()
        {
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Reclutamiento.MisTallas")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null &&
                                // Data Tallas  RL///////////////////
                                rows["TddlCodPrens"].ToString().Length != 0 && rows["TddlCodPrens"].ToString() != null &&
                                rows["TddlDetTallas"].ToString().Length != 0 && rows["TddlDetTallas"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                // Data Tallas  RL///////////////////
                                string TddlCodPrens = rows["TddlCodPrens"].ToString();
                                string TddlDetTallas = rows["TddlDetTallas"].ToString();
                                string Url = rows["Url"].ToString();
                                string Maquina = rows["Maquina"].ToString();

                                try
                                {
                                    string database = "";
                                    if (Url.ToLower().Contains("ora"))
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
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();

                                    //login
                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);

                                    //Process: Mis Tallas ///////////////////////////////////////////////////////////////////
                                    selenium.Click("//a[contains(text(),'Mis Tallas')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Tallas", true, file);

                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodPren')]",TddlCodPrens);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlDetTalla')]",TddlDetTallas);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Datos Tallas", true, file);

                                    selenium.Click("//button[contains(@id,'btnGuardar')]");

                                    if(selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_grvBiEmtal_ctl02_LinkButton1']"))
                                    {
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Existe Datos Tallas", true, file);

                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_grvBiEmtal_ctl02_LinkButton1']");
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Elimina Datos Tallas", true, file);

                                    }
                                    else
                                    {
                                        Thread.Sleep(500);
                                        selenium.Screenshot("No existe Datos Tallas", true, file);

                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se inserto la talla");
                                    }
                                    //////
                                    ConverWordToPDF(file, database);
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
        public void NuevoRegistro()
        {
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Reclutamiento.NuevoRegistro")
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
                                rows["TipDocum"].ToString().Length != 0 && rows["TipDocum"].ToString() != null &&
                                rows["NumDocum"].ToString().Length != 0 && rows["NumDocum"].ToString() != null &&
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["Apellido"].ToString().Length != 0 && rows["Apellido"].ToString() != null &&
                                rows["Email"].ToString().Length != 0 && rows["Email"].ToString() != null &&
                                rows["Contraseña"].ToString().Length != 0 && rows["Contraseña"].ToString() != null &&
                                rows["Pregunta"].ToString().Length != 0 && rows["Pregunta"].ToString() != null &&
                                rows["Respuesta"].ToString().Length != 0 && rows["Respuesta"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null &&
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                string TipDocum = rows["TipDocum"].ToString();
                                string NumDocum = rows["NumDocum"].ToString();
                                string Nombre = rows["Nombre"].ToString();
                                string Apellido = rows["Apellido"].ToString();
                                string Email = rows["Email"].ToString();
                                string Contraseña = rows["Contraseña"].ToString();
                                string Pregunta = rows["Pregunta"].ToString();
                                string Respuesta = rows["Respuesta"].ToString();
                                string Url = rows["Url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string User = rows["User"].ToString();


                                try
                                {
                                    string database = "";
                                    if (Url.ToLower().Contains("ora"))
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
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;

                                    string Borrar1Tabla = $"Delete from RW_WREGI where NUM_IDEN ={NumDocum}";
                                    db.UpdateDeleteInsert(Borrar1Tabla, database, User);

                                    List<string> errorMessagesMetodo = new List<string>();
                                    var options = new ChromeOptions();
                                    options.AddArgument("-no-sandbox");
                                    OpenQA.Selenium.Chrome.ChromeDriver driver = new ChromeDriver(@"C:\deployment\", options, TimeSpan.FromSeconds(3600));
                                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3600);

                                    driver.Manage().Window.Maximize();
                                    driver.Navigate().GoToUrl(Url);
                                    Thread.Sleep(500);
                                    Screenshot(modulo, true, file);

                                    driver.FindElement(By.XPath("//a[contains(text(),'Registrarse')]")).Click();


                                    SelectElement TipDocums = new SelectElement(driver.FindElement(By.XPath("//select[contains(@id,'ddlTipDocu')]")));
                                    TipDocums.SelectByText(TipDocum);

                                    driver.FindElement(By.XPath("//input[contains(@id,'txtCodEmpl1')]")).SendKeys(NumDocum);
                                    Thread.Sleep(500);
                                    driver.FindElement(By.XPath("//input[contains(@id,'txtNomEmpl')]")).SendKeys(Nombre);
                                    driver.FindElement(By.XPath("//input[contains(@id,'txtNomEmpl')]")).SendKeys(Nombre);
                                    Thread.Sleep(500);
                                    driver.FindElement(By.XPath("//input[contains(@id,'txtApeEmpl')]")).SendKeys(Apellido);
                                    driver.FindElement(By.XPath("//input[contains(@id,'txtBoxMail')]")).SendKeys(Email);
                                    driver.FindElement(By.XPath("//input[contains(@id,'txtBoxMailC')]")).SendKeys(Email);
                                    driver.FindElement(By.XPath("//input[contains(@id,'txtBoxMailC')]")).SendKeys(Email);
                                    driver.FindElement(By.XPath("//input[contains(@id,'txtPasUsua1')]")).SendKeys(Contraseña);
                                    driver.FindElement(By.XPath("//input[contains(@id,'txtPasUsuaC')]")).SendKeys(Contraseña);

                                    SelectElement Preguntas = new SelectElement(driver.FindElement(By.XPath("//select[contains(@id,'ddlPrePass1')]")));
                                    Preguntas.SelectByText(Pregunta);

                                    driver.FindElement(By.XPath("//input[contains(@id,'txtResPass1')]")).SendKeys(Respuesta);

                                    Screenshot("Datos Registro", true, file);


                                    driver.FindElement(By.XPath("//input[contains(@id,'btnGuardarRegistro')]")).Click();

                                    Thread.Sleep(500);

                                    Screenshot("Aceptación de Términos y Condiciones", true, file);


                                    driver.FindElement(By.XPath("//input[contains(@id,'btnAcepto')]")).Click();

                                    Screenshot("Aceptación del Habeas Data", true, file);


                                    driver.FindElement(By.XPath("//input[contains(@id,'btnSiH')]")).Click();

                                   Screenshot("Registro Exitoso", true, file);

                                    //////
                                    ConverWordToPDF(file, database);
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
        public void Publicaciones()
        {
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Reclutamiento.Publicaciones")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null &&
                                // Data Publicaciones  RL/////////////////////
                                rows["PuPubl"].ToString().Length != 0 && rows["PuPubl"].ToString() != null &&
                                rows["PuEdi"].ToString().Length != 0 && rows["PuEdi"].ToString() != null &&
                                rows["PuddlPUBs"].ToString().Length != 0 && rows["PuddlPUBs"].ToString() != null &&
                                rows["PutxtFecha"].ToString().Length != 0 && rows["PutxtFecha"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                // Data Publicaciones  RL/////////////////////
                                string PuPubl = rows["PuPubl"].ToString();
                                string PuEdi = rows["PuEdi"].ToString();
                                string PuddlPUBs = rows["PuddlPUBs"].ToString();
                                string PutxtFecha = rows["PutxtFecha"].ToString();
                                string Url = rows["Url"].ToString();
                                string Maquina = rows["Maquina"].ToString();

                                try
                                {
                                    string database = "";
                                    if (Url.ToLower().Contains("ora"))
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
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();
                                    
                                    //login
                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);

                                    //Process: PUblicaciones ///////////////////////////////////////////////////////////////////
                                    selenium.Click("//a[contains(text(),'Publicaciones')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Publicaciones", true, file);

                                    selenium.SendKeys("//input[contains(@id,'txtTit_Publ')]",PuPubl);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtEDI_PUBL')]",PuEdi);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlPUB_CLAS')]",PuddlPUBs);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'KCtrlFEC_PUBL_txtFecha')]",PutxtFecha);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Datos Publicaciones", true, file);

                                    selenium.Click("//button[contains(@id,'btnGuardar')]");

                                    if(selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_dgPubli_ctl03_LinkButton2']"))
                                    {
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Existe Datos Publicaciones", true, file);

                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_dgPubli_ctl03_LinkButton2']");
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Elimina Datos Publicaciones", true, file);

                                    }
                                    else
                                    {
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("No existe Datos Publicaciones", true, file);

                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se inserto la publcacion");
                                    }
                                    //////
                                    ConverWordToPDF(file, database);
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
        public void Trayectoria()
        {
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Reclutamiento.Trayectoria")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null &&
                                // Data Trayectoria RL/////////////////////
                                rows["TrtxtNomInstDtrinI"].ToString().Length != 0 && rows["TrtxtNomInstDtrinI"].ToString() != null &&
                                rows["TrtxtAutOriaI"].ToString().Length != 0 && rows["TrtxtAutOriaI"].ToString() != null &&
                                rows["TrtxtNomProyI"].ToString().Length != 0 && rows["TrtxtNomProyI"].ToString() != null &&
                                rows["TrtxtFecha"].ToString().Length != 0 && rows["TrtxtFecha"].ToString() != null &&
                                rows["TrtxtFechaFin"].ToString().Length != 0 && rows["TrtxtFechaFin"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                // Data Trayectoria RL/////////////////////
                                string TrtxtNomInstDtrinI = rows["TrtxtNomInstDtrinI"].ToString();
                                string TrtxtAutOriaI = rows["TrtxtAutOriaI"].ToString();
                                string TrtxtNomProyI = rows["TrtxtNomProyI"].ToString();
                                string TrtxtFecha = rows["TrtxtFecha"].ToString();
                                string TrtxtFechaFin = rows["TrtxtFechaFin"].ToString();
                                string Url = rows["Url"].ToString();
                                string Maquina = rows["Maquina"].ToString();

                                try
                                {
                                    string database = string.Empty;
                                    if (Url.ToLower().Contains("ora"))
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
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();

                                    //login
                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);

                                    //Process: Trayectoria ///////////////////////////////////////////////////////////////////
                                    selenium.Click("//a[contains(text(),'Trayectoria Investigativa')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Trayectoria", true, file);
                                    selenium.Click("//input[contains(@id,'LinkButton1')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel1_btnNuevoDtrinI']");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtNomInstDtrinI')]",TrtxtNomInstDtrinI);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtAutOriaI')]",TrtxtAutOriaI);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'txtNomProyI')]",TrtxtNomProyI);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'KCtrlFecInicI_txtFecha')]",TrtxtFecha);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'KCtrlFecFinaI_txtFecha')]",TrtxtFechaFin);             
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Datos Trayectoria", true, file);

                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel1_btnGuardarDtrinI']");

                                    if(selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel1_dgrBiDtrinI_ctl03_LinkButton2']"))
                                    {
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Existe Datos Trayectoria", true, file);

                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_TabContainer1_TabPanel1_dgrBiDtrinI_ctl03_LinkButton2']");
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Elimina Datos Trayectoria", true, file);

                                    }
                                    else
                                    {
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("no existe Datos Trayectoria", true, file);

                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se inserto la trayectoria");
                                    }
                                    //////
                                    ConverWordToPDF(file, database);
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
