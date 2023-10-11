using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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


namespace Web_Kactus_Test_V2
{
    /// <summary>
    /// Descripción resumida de CodedUITest1
    /// </summary>
    [TestClass]
    public class Reclutamiento : FuncionesVitales
    {
        string app = "Reclutamiento";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();
        public Reclutamiento()
        {
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Reclutamiento.DatosBasicos")
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


                            ////

                            if (
                                //Datos Login
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RLPassword"].ToString().Length != 0 && rows["RLPassword"].ToString() != null &&
                                // Data Trayectoria RL/////////////////////
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["Perfil"].ToString().Length != 0 && rows["Perfil"].ToString() != null &&
                                rows["URL"].ToString().Length != 0 && rows["URL"].ToString() != null

                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RLPassword"].ToString();
                                // Data Trayectoria RL/////////////////////
                                string Nombre = rows["Nombre"].ToString();
                                string Perfil = rows["Perfil"].ToString();
                                string Url = rows["URL"].ToString();
                                string Area = rows["AreaInteres"].ToString();

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
                                    string user = "";
                                    string database = "";
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestAuto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestoraAuto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                    }

                                    if (Url.ToLower() == "http://dwtfsk:8088/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                    }
                                    if (Url.ToLower() == "http://dwtfsk:8087/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();

                                    //PARAMETRIZACION PREVIA
                                    if (database == "SQL")
                                    {
                                        string deletedocu = $"delete from RL_AREMP where cod_empl='845267621' and cod_empr='9'";
                                        db.UpdateDeleteInsert(deletedocu, database, user);

                                    }
                                    else
                                    {
                                        string deletedocu = $"delete from RL_AREMP where cod_empl='1011121314' and cod_empr='421'";
                                        db.UpdateDeleteInsert(deletedocu, database, user);

                                    }
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
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAreInte']", Area);
                                    Thread.Sleep(2000);
                                    selenium.ScrollTo("0", "900");
                                    selenium.Screenshot("Registro Area Experiencia", true, file);
                                    //ACTUALIZAR
                                    selenium.Click("//a[@id='btnActualizar']");
                                    Thread.Sleep(2000);
                                
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Actualización exitosa", true, file);
                                        Thread.Sleep(2000);

                                    
                                    

                                    //////
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Reclutamiento.Documentos")
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


                            ////

                            if (
                                //Datos Login
                                rows["RLUser"].ToString().Length != 0 && rows["RLUser"].ToString() != null &&
                                rows["RLPass"].ToString().Length != 0 && rows["RLPass"].ToString() != null &&
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["URL"].ToString().Length != 0 && rows["URL"].ToString() != null &&
                                // Data Documentos  RL/////////////////////
                                rows["TipoDocumento"].ToString().Length != 0 && rows["TipoDocumento"].ToString() != null &&
                                rows["NumeroDocumento"].ToString().Length != 0 && rows["NumeroDocumento"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["FechaExpedicion"].ToString().Length != 0 && rows["FechaExpedicion"].ToString() != null &&
                                rows["FechaVencimiento"].ToString().Length != 0 && rows["FechaVencimiento"].ToString() != null &&
                                rows["Observaciones"].ToString().Length != 0 && rows["Observaciones"].ToString() != null

                                )
                            {
                                //LOGIN
                                string RlUser = rows["RLUser"].ToString();
                                string RlPass = rows["RLPass"].ToString();
                                string Url = rows["URL"].ToString();
                                string User = rows["User"].ToString();
                                // Data Documentos  RL/////////////////////
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string NumeroDocumento = rows["NumeroDocumento"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string FechaExpedicion = rows["FechaExpedicion"].ToString();
                                string FechaVencimiento = rows["FechaVencimiento"].ToString();
                                string Observaciones = rows["Observaciones"].ToString();

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

                                    string database = "";
                                    if (Url.ToLower() == "http://dwtfskscm/reclutamientotestauto/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    if (Url.ToLower() == "http://dwtfskscm/reclutamientotestORAauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (Url.ToLower() == "http://dwtfsk:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    if (Url.ToLower() == "http://dwtfsk:8087/".ToLower())
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
                                    //INICIO PRUEBAS
                                    //PARAMETRIZACIONES

                                    //PARAMETRIZACION PREVIA
                                    if (database == "SQL")
                                    {
                                        string deletedocu = $"delete from rl_empdo where cod_empl='845267621' and cod_empr='9' and num_docu='789456'";
                                        db.UpdateDeleteInsert(deletedocu, database, User);

                                    }
                                    else if (database == "ORA")
                                    {
                                        string deletedocu = $"delete from rl_empdo where cod_empl='1011121314' and cod_empr='421' and num_docu='789456'";
                                        db.UpdateDeleteInsert(deletedocu, database, User);

                                    }
                                    //PRUEBA
                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);

                                    //DOCUMENTOS
                                    selenium.Click("//a[contains(.,'Documentos')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Mis documentos", true, file);
                                    //NUEVO
                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nuevo Documento", true, file);
                                    //DATOS
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomDocu']", TipoDocumento);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Tipo Documento", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNumDocu']", NumeroDocumento);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Numero Documento", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrDivipPais_txtDivPoli']", Ciudad);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Ciudad", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecExpe_txtFecha']", FechaExpedicion);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Fecha Expedicion", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecVenc_txtFecha']", FechaVencimiento);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Fecha Vencimiento", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_txtObsErva_txtTexto']", Observaciones);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(3000);
                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro Guardado", true, file);
                                    Thread.Sleep(3000);

                                    //EDITAR REGISTRO
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_grvRlEmpdo_ctl02_LinkButton1']/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Editar Documento Guardado", true, file);
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_fuArcAdju']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_txtObsErva_txtTexto']", "EDICION DE OBSERVACIONES");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Observaciones Editado", true, file);
                                    Thread.Sleep(3000);
                                    //ACTUALIZAR
                                    selenium.Click("//a[@id='btnActualizar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Edicion Documento Exitosa", true, file);
                                    Thread.Sleep(3000);
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Reclutamiento.EducacionFormal")
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


                            ////

                            if (
                                //Datos Login
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null &&
                                // Data Educacion Formal  RL///////////
                                rows["Modalidad"].ToString().Length != 0 && rows["Modalidad"].ToString() != null &&
                                rows["NombreEstudios"].ToString().Length != 0 && rows["NombreEstudios"].ToString() != null &&
                                rows["NombreEspecifico"].ToString().Length != 0 && rows["NombreEspecifico"].ToString() != null &&
                                rows["Institucion"].ToString().Length != 0 && rows["Institucion"].ToString() != null &&
                                rows["TiempoEstudio"].ToString().Length != 0 && rows["TiempoEstudio"].ToString() != null &&
                                rows["UnidadTiempo"].ToString().Length != 0 && rows["UnidadTiempo"].ToString() != null &&
                                rows["FechaInicio"].ToString().Length != 0 && rows["FechaInicio"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["TarjetaProfesional"].ToString().Length != 0 && rows["TarjetaProfesional"].ToString() != null &&
                                rows["Promedio"].ToString().Length != 0 && rows["Promedio"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["Metodologia"].ToString().Length != 0 && rows["Metodologia"].ToString() != null



                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                string Url = rows["Url"].ToString();
                                // Data Educacion Formal  RL///////////
                                string Modalidad = rows["Modalidad"].ToString();
                                string NombreEstudios = rows["NombreEstudios"].ToString();
                                string NombreEspecifico = rows["NombreEspecifico"].ToString();
                                string Institucion = rows["Institucion"].ToString();
                                string TiempoEstudio = rows["TiempoEstudio"].ToString();
                                string UnidadTiempo = rows["UnidadTiempo"].ToString();
                                string FechaInicio = rows["FechaInicio"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string TarjetaProfesional = rows["TarjetaProfesional"].ToString();
                                string Promedio = rows["Promedio"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string Metodologia = rows["Metodologia"].ToString();

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
                                    string user = "";
                                    string database = "";
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestAuto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestoraAuto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                    }

                                    if (Url.ToLower() == "http://dwtfsk:8088/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                    }
                                    if (Url.ToLower() == "http://dwtfsk:8087/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }

                                    //PARAMETRIZACION PREVIA
                                    if (database == "SQL")
                                    {
                                        string deletedocu = $"delete FROM RL_DTPEF WHERE RMT_EDFO='281'";
                                        db.UpdateDeleteInsert(deletedocu, database, user);
                                        string delete = $"delete from rl_edfor where cod_empl='845267621' and cod_empr='9' and mat_prof='123456789'";
                                        db.UpdateDeleteInsert(delete, database, user);

                                    }
                                    else
                                    {
                                        string deletedocu = $"delete FROM RL_DTPEF WHERE RMT_EDFO='37'";
                                        db.UpdateDeleteInsert(deletedocu, database, user);
                                        string delete = $"delete from rl_edfor where cod_empl='1011121314' and cod_empr='421' and mat_prof='123456789'";
                                        db.UpdateDeleteInsert(delete, database, user);

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
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);
                                    //EDUCACION FORMAL
                                    selenium.Click("//a[contains(.,'Edu. Formal')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Educación Formal Reclutamiento", true, file);
                                    //NUEVO REGISTRO
                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nuevo Registro Educación Formal", true, file);
                                    //LLENADO DE DATOS
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomModi']", Modalidad);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Modalidad", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomEstu']", NombreEstudios);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nombre estudios", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomEspp']", NombreEspecifico);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nombre Específico", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_ddlNomInst']", Institucion);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Institucion", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtTieEstu']", TiempoEstudio);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Tiempo Estudio", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlUniTiem']", UnidadTiempo);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Unidad Estudio", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", FechaInicio);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Fecha Inicio", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecTerm_txtFecha']", FechaFinal);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Fecha Final", true, file);
                                    Thread.Sleep(3000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_chkTapEntr']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMatProf']", TarjetaProfesional);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Tarjeta Profesional", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtProCarr']", Promedio);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Promedio", true, file);
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrDivipPais_txtDivPoli']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrDivipPais_txtDivPoli']", Ciudad);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Ciudad", true, file);
                                    Thread.Sleep(3000);
                                    if (database == "SQL")
                                    {
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodItem']", Metodologia);
                                        Thread.Sleep(3000);
                                        selenium.Screenshot("Metodologia", true, file);
                                        Thread.Sleep(3000);
                                    }
                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Educacion Formal Registrada", true, file);
                                    Thread.Sleep(3000);
                                    DateTime dateAndTime = DateTime.Now;
                                    string fecha = dateAndTime.ToString("yyyyMMdd");
                                    string actualizacion = NombreEspecifico + fecha;
                                    //EDICION DETALLE
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_grvRlEdfor_ctl02_LinkButton1']/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Educacion Formal Editar", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomEspp']", actualizacion);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Campo Nombre Especifico Actualizado", true, file);
                                    Thread.Sleep(3000);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrDivipPais_txtDivPoli']", Ciudad);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Ciudad", true, file);
                                    //ACTUALIZAR
                                    selenium.Click("//a[@id='btnActualizar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro Actualizado", true, file);
                                    Thread.Sleep(3000);
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Reclutamiento.EducacionNoFormal")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null
                                // Data Educacion no Formal  RL////////

                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                // Data Educacion no Formal  RL////////
                                string Modalidad = rows["Modalidad"].ToString();
                                string NombreEstudios = rows["NombreEstudios"].ToString();
                                string NombreEspecifico = rows["NombreEspecifico"].ToString();
                                string NombreInstitucion = rows["NombreInstitucion"].ToString();
                                string FechaInicio = rows["FechaInicio"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string TiempoEstudio = rows["TiempoEstudio"].ToString();
                                string UnidadTiempoEstudio = rows["UnidadTiempoEstudio"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string Url = rows["Url"].ToString();

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
                                    string user = "";
                                    string database = "";
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestAuto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestoraAuto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                    }

                                    if (Url.ToLower() == "http://dwtfsk:8088/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                    }
                                    if (Url.ToLower() == "http://dwtfsk:8087/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    DateTime dateAndTime = DateTime.Now;
                                    string datetime = dateAndTime.ToString("ddMMyyyyHHmmss");
                                    string UpdateData = "TEST_" + datetime;
                                    // Data Basicos RL//////
                                    string NomAspirante = UpdateData;
                                    ///////////////////////////////////////
                                    if (database == "SQL")
                                    {
                                        string deletedocu = $"delete FROM RL_DTPNF WHERE cod_empr='9' AND RMT_EDNF='289'";
                                        db.UpdateDeleteInsert(deletedocu, database, user);
                                        string delete = $"delete from rl_ednfo WHERE cod_empr='9' AND cod_empl='845267621' AND mod_acad='CU'";
                                        db.UpdateDeleteInsert(delete, database, user);

                                    }
                                    else
                                    {
                                        string deletedocu = $"delete FROM RL_DTPNF WHERE cod_empr='421' AND RMT_EDNF='54'";
                                        db.UpdateDeleteInsert(deletedocu, database, user);
                                        string delete = $"delete from rl_ednfo WHERE cod_empr='421' AND cod_empl='1011121314' AND mod_acad='CU'";
                                        db.UpdateDeleteInsert(delete, database, user);


                                    }
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();

                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);

                                    //Process: Educacion no Fromal////////////////////////////////////////////////////////
                                    selenium.Click("//a[contains(.,'Edu. No Formal')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Edu. No Formal", true, file);
                                    //NUEVO
                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nuevo Registro", true, file);
                                    //DIlIGENCIAR DATOS
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomModi']", Modalidad);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Modalidad", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomEstu']", NombreEstudios);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nombre Estudios", true, file);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomEspe']", NombreEspecifico);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nombre Especifico", true, file);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomInst']", NombreInstitucion);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nombre Institucion", true, file);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", FechaInicio);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Fecha Inicio", true, file);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecTerm_txtFecha']", FechaFinal);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Fecha Final", true, file);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtTieEstu']", TiempoEstudio);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Tiempo Estudio", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlUniTiem']", UnidadTiempoEstudio);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Unidad Tiempo Estudio", true, file);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrDivipPais_txtDivPoli']", Ciudad);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Ciudad", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro Guardado", true, file);

                                    //ACTUALIZAR REGISTRO
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_grvRlEdnfo_ctl03_LinkButton1']/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro a Editar", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomEspe']", NombreEspecifico + datetime);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Campo Editado", true, file);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrDivipPais_txtDivPoli']", Ciudad);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Ciudad", true, file);
                                    selenium.Click("//a[@id='btnActualizar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro Editado", true, file);
                                    Thread.Sleep(3000);
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Reclutamiento.ExperienciaLaboral")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null

                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                string Url = rows["Url"].ToString();
                                // Data Experiencia Laboral //
                                string Empresa = rows["Empresa"].ToString();
                                string Direccion = rows["Direccion"].ToString();
                                string Telefono = rows["Telefono"].ToString();
                                string TipoEmpresa = rows["TipoEmpresa"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string Cargo = rows["Cargo"].ToString();
                                string Dedicado = rows["Dedicado"].ToString();
                                string TipoContrato = rows["TipoContrato"].ToString();
                                string ManejaPersonal = rows["ManejaPersonal"].ToString();
                                string Actividad = rows["Actividad"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string Jefe = rows["Jefe"].ToString();
                                string CargoJefe = rows["CargoJefe"].ToString();
                                string Funciones = rows["Funciones"].ToString();
                                string AreaExperiencia = rows["AreaExperiencia"].ToString();

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
                                    string user = "";
                                    string database = "";
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestAuto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestoraAuto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                    }

                                    if (Url.ToLower() == "http://dwtfsk:8088/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                    }
                                    if (Url.ToLower() == "http://dwtfsk:8087/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }

                                    if (database == "SQL")
                                    {
                                        string deletedocu = $"delete from RL_hvext where cod_empl='845267621' and cod_empr='9'";
                                        db.UpdateDeleteInsert(deletedocu, database, user);

                                    }
                                    else
                                    {
                                        string deletedocu = $"delete from RL_hvext where cod_empl='1011121314' and cod_empr='421'";
                                        db.UpdateDeleteInsert(deletedocu, database, user);

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
                                    selenium.Click("//a[contains(.,'Exp Laboral')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Experiencia laboral", true, file);
                                    //nuevo
                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Experiencia laboral", true, file);
                                    //Diligenciar datos
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomEmpr']", Empresa);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtDirEmpr']", Direccion);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtTelEmpr']", Telefono);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlTipEmpr']", TipoEmpresa);
                                    Thread.Sleep(2000);
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_UpdatePanel1']/div/div/span/label");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecIngr_txtFecha']", FechaInicial);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtCarDese']", Cargo);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDedIcac']", Dedicado);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlTipCont']", TipoContrato);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlManPers']", ManejaPersonal);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodActi']", Actividad);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrDivipPais_txtDivPoli']", Ciudad);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtJefInme']", Jefe);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtCarJefe']", CargoJefe);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_txtFunReal_txtTexto']", Funciones);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAreExpe']", AreaExperiencia);
                                    Thread.Sleep(2000);
                                    if (selenium.ExistControl("//div[4]/div/button"))
                                    {
                                        Thread.Sleep(2000);
                                        selenium.Click("//div[4]/div/button");
                                    }
                                    selenium.Screenshot("Datos", true, file);
                                    Thread.Sleep(2000);
                                    
                                    //Guardar
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro Guardado", true, file);
                                    if (database == "ORA")
                                    {
                                        if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                        {
                                            selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                        }
                                    }
                                    //Proyectos
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnNoProLog']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro Guardado", true, file);
                                    if (selenium.ExistControl("//div[4]/div/button"))
                                    {
                                        Thread.Sleep(2000);
                                        selenium.Click("//div[4]/div/button");
                                    }

                                    //EDITAR REGISTRO
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_grvRlHvext_ctl02_LinkButton1']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro Editar", true, file);
                                    //Editar campo empresa
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomEmpr']", Empresa + " Edicion");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro Editado", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrDivipPais_txtDivPoli']", Ciudad);
                                    Thread.Sleep(2000);
                                    //Actualizar
                                    selenium.Click("//*[@id='btnActualizar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro Editado", true, file);
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Reclutamiento.Familiares")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null
                                // Data Familiares  RL/////////////////

                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                string Url = rows["Url"].ToString();
                                // Data Familiares  RL/////////////////
                                string IdentificacionFamiliar = rows["IdentificacionFamiliar"].ToString();
                                string NombreFamiliar = rows["NombreFamiliar"].ToString();
                                string NombreFamiliar2 = rows["NombreFamiliar2"].ToString();
                                string ApellidosFamiliar = rows["ApellidosFamiliar"].ToString();
                                string ApellidosFamiliar2 = rows["ApellidosFamiliar2"].ToString();
                                string FechaNacimiento = rows["FechaNacimiento"].ToString();
                                string GrupoSanguineo = rows["GrupoSanguineo"].ToString();
                                string FactorSanguineo = rows["FactorSanguineo"].ToString();
                                string EstadoCivil = rows["EstadoCivil"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();

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
                                    string user = "";
                                    string database = "";
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestAuto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestoraAuto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                    }

                                    if (Url.ToLower() == "http://dwtfsk:8088/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                    }
                                    if (Url.ToLower() == "http://dwtfsk:8087/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }

                                    if (database == "SQL")
                                    {

                                        string delete = $"delete FROM rl_famil  WHERE cod_empr='9' and cod_fami='789456' and cod_empl='845267621'";
                                        db.UpdateDeleteInsert(delete, database, user);

                                    }
                                    else
                                    {
                                        string delete = $"delete FROM rl_famil  WHERE cod_empr='421' and cod_fami='789456' and cod_empl='1011121314'";
                                        db.UpdateDeleteInsert(delete, database, user);

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
                                    selenium.Click("//a[contains(.,'Familiares')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Familiares", true, file);

                                    //NUEVO
                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Nuevo Registro", true, file);
                                    //DILIGENCIAR DATOS
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtCodFami']", IdentificacionFamiliar);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomFami1']", NombreFamiliar);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomFami2']", NombreFamiliar2);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtApeFami1']", ApellidosFamiliar);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtApeFami2']", ApellidosFamiliar2);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecNaci_txtFecha']", FechaNacimiento);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecNaci_txtFecha']", FechaNacimiento);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlGruSang']", GrupoSanguineo);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlFacSang']", FactorSanguineo);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlEstCivi']", EstadoCivil);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrDivip_txtDivPoli']", Ciudad);
                                    Thread.Sleep(2000);
                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro Exitoso", true, file);
                                    //ACTUALIZAR REGUISTRO FAMILIARES
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgEdFamil_ctl02_LinkButton1']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro a Editar", true, file);
                                    //CAMPO A EDITAR
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtApeFami1']", ApellidosFamiliar + "Editado");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrDivip_txtDivPoli']", Ciudad);
                                    Thread.Sleep(2000);
                                    //ACTUALIZAR
                                    selenium.Click("//a[@id='btnActualizar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro Editado", true, file);
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Reclutamiento.Idiomas")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null
                                // Data Idiomas  RL///////////////////

                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                string Url = rows["Url"].ToString();
                                // Data Idiomas  RL//////////////////
                                string Idioma = rows["Idioma"].ToString();
                                string Habla = rows["Habla"].ToString();
                                string Lee = rows["Lee"].ToString();
                                string Escribe = rows["Escribe"].ToString();
                                string Observaciones = rows["Observaciones"].ToString();
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
                                    string user = "";
                                    string database = "";
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestAuto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestoraAuto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                    }

                                    if (Url.ToLower() == "http://dwtfsk:8088/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                    }
                                    if (Url.ToLower() == "http://dwtfsk:8087/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();
                                    //LOGIN
                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);
                                    //Process: Idiomas ///////////////////////////////////////////////////////////////////
                                    selenium.Click("//a[contains(.,'Idiomas')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Idiomas", true, file);
                                    //DILIGENCIAR IDIOMAS
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomIdio']", Idioma);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Idioma", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlHabIdio']", Habla);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Habla", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlLeeIdio']", Lee);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Lee", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlEscIdio']", Escribe);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Escribe", true, file);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtObsErva']", Observaciones);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro Guardado", true, file);
                                    //ELIMINAR REGISTRO
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_grvRlEmidi_ctl02_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro Eliminado", true, file);
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Reclutamiento.MisTallas")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null

                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                // Data Tallas  RL///////////////////
                                string Talla = rows["Talla"].ToString();
                                string Prenda = rows["Prenda"].ToString();
                                string Url = rows["Url"].ToString();
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
                                    string user = "";
                                    string database = "";
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestAuto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestoraAuto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                    }

                                    if (Url.ToLower() == "http://dwtfsk:8088/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                    }
                                    if (Url.ToLower() == "http://dwtfsk:8087/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
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
                                    selenium.Click("//a[contains(.,'Mis Tallas')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tallas", true, file);
                                    //DILIGENCIAR TALLAS
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodPren']", Prenda);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Prenda", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDetTalla']", Talla);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Talla", true, file);
                                    //Guardar
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro Exitoso", true, file);
                                    //ELIMINAR
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_grvBiEmtal_ctl02_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_btnCerrar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro Eliminado", true, file);
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Reclutamiento.NuevoRegistro")
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
                                rows["TipoDocumento"].ToString().Length != 0 && rows["TipoDocumento"].ToString() != null &&
                                rows["Documento"].ToString().Length != 0 && rows["Documento"].ToString() != null &&
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["Apellidos"].ToString().Length != 0 && rows["Apellidos"].ToString() != null &&
                                rows["Correo"].ToString().Length != 0 && rows["Correo"].ToString() != null &&
                                rows["Clave"].ToString().Length != 0 && rows["Clave"].ToString() != null &&
                                rows["Pregunta"].ToString().Length != 0 && rows["Pregunta"].ToString() != null &&
                                rows["Respuesta"].ToString().Length != 0 && rows["Respuesta"].ToString() != null &&
                                rows["Url"].ToString().Length != 0 && rows["Url"].ToString() != null

                                )
                            {
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string Documento = rows["Documento"].ToString();
                                string Nombre = rows["Nombre"].ToString();
                                string Apellidos = rows["Apellidos"].ToString();
                                string Correo = rows["Correo"].ToString();
                                string Clave = rows["Clave"].ToString();
                                string Pregunta = rows["Pregunta"].ToString();
                                string Respuesta = rows["Respuesta"].ToString();
                                string Url = rows["Url"].ToString();

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
                                    string user = "";
                                    string database = "";
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestAuto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestoraAuto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                    }

                                    if (Url.ToLower() == "http://dwtfsk:8088/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                    }
                                    if (Url.ToLower() == "http://dwtfsk:8087/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;

                                    string Borrar1Tabla = $"Delete from RW_WREGI where NUM_IDEN ='{Documento}'";
                                    db.UpdateDeleteInsert(Borrar1Tabla, database, user);
                                    List<string> errorMessagesMetodo = new List<string>();
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

                                    //login
                                    var options = new ChromeOptions();
                                    ChromeDriver driver = new ChromeDriver(@"C:\deployment\", options, TimeSpan.FromSeconds(240));
                                    driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(240);
                                    driver.Manage().Window.Maximize();
                                    driver.Navigate().GoToUrl(Url);
                                    //REGISTRARSE
                                    Thread.Sleep(7000);
                                    driver.FindElement(By.XPath("//a[contains(text(),'Registrarse')]")).Click();
                                    Thread.Sleep(7000);
                                    //DILIGENCIAR NUEVO REGISTRO
                                    SelectElement select = new SelectElement(driver.FindElement(By.XPath("//select[@id='ddlTipDocu']")));
                                    select.SelectByText(TipoDocumento);
                                    Thread.Sleep(7000);
                                    driver.FindElement(By.XPath("//input[@id='txtCodEmpl1']")).Clear();
                                    driver.FindElement(By.XPath("//input[@id='txtCodEmpl1']")).SendKeys(Documento);
                                    Thread.Sleep(7000);
                                    SendKeys.SendWait("{ENTER}");
                                    driver.FindElement(By.XPath("//input[@id='txtNomEmpl']")).Clear();
                                    driver.FindElement(By.XPath("//input[@id='txtNomEmpl']")).SendKeys(Nombre);
                                    Thread.Sleep(7000);
                                    SendKeys.SendWait("{ENTER}");
                                    driver.FindElement(By.XPath("//input[@id='txtApeEmpl']")).Clear();
                                    driver.FindElement(By.XPath("//input[@id='txtApeEmpl']")).SendKeys(Apellidos);
                                    Thread.Sleep(7000);
                                    SendKeys.SendWait("{ENTER}");
                                    driver.FindElement(By.XPath("//input[@id='txtBoxMail']")).Clear();
                                    driver.FindElement(By.XPath("//input[@id='txtBoxMail']")).SendKeys(Correo);
                                    Thread.Sleep(7000);
                                    SendKeys.SendWait("{ENTER}");
                                    driver.FindElement(By.XPath("//input[@id='txtBoxMailC']")).Clear();
                                    driver.FindElement(By.XPath("//input[@id='txtBoxMailC']")).SendKeys(Correo);
                                    Thread.Sleep(7000);
                                    SendKeys.SendWait("{ENTER}");
                                    driver.FindElement(By.XPath("//input[@id='txtPasUsua1']")).Clear();
                                    driver.FindElement(By.XPath("//input[@id='txtPasUsua1']")).SendKeys(Clave);
                                    Thread.Sleep(7000);
                                    SendKeys.SendWait("{ENTER}");
                                    driver.FindElement(By.XPath("//input[@id='txtPasUsuaC']")).Clear();
                                    driver.FindElement(By.XPath("//input[@id='txtPasUsuaC']")).SendKeys(Clave);
                                    Thread.Sleep(7000);
                                    SelectElement select1 = new SelectElement(driver.FindElement(By.XPath("//select[@id='ddlPrePass1']")));
                                    select1.SelectByText(Pregunta);
                                    Thread.Sleep(7000);
                                    driver.FindElement(By.XPath("//input[@id='txtResPass1']")).Clear();
                                    driver.FindElement(By.XPath("//input[@id='txtResPass1']")).SendKeys(Respuesta);
                                    Thread.Sleep(7000);
                                    Screenshot("Datos", true, file);
                                    //Guardar Registro
                                    driver.FindElement(By.XPath("//input[@id='btnGuardarRegistro']")).Click();
                                    Thread.Sleep(7000);
                                    Screenshot("Guardar", true, file);
                                    //Terminos y condiciones
                                    driver.FindElement(By.XPath("//*[@id='btnAcepto']")).Click();
                                    Thread.Sleep(7000);
                                    driver.FindElement(By.XPath("//*[@id='btnSiH']")).Click();
                                    Thread.Sleep(10000);
                                    Screenshot("Registro Exitoso", true, file);
                                    driver.Close();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Reclutamiento.Publicaciones")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null

                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                // Data Publicaciones  RL/////////////////////
                                string Titulo = rows["Titulo"].ToString();
                                string Editorial = rows["Editorial"].ToString();
                                string ISBN = rows["ISBN"].ToString();
                                string TipoPublicacion = rows["TipoPublicacion"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string Url = rows["Url"].ToString();

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
                                    string user = "";
                                    string database = "";
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestAuto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestoraAuto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                    }

                                    if (Url.ToLower() == "http://dwtfsk:8088/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                    }
                                    if (Url.ToLower() == "http://dwtfsk:8087/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
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
                                    //Publicaciones
                                    selenium.Click("//a[contains(.,'Publicaciones')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Publicaciones", true, file);
                                    //Diligenciar datos
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtTit_Publ']", Titulo);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Titulo", true, file);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtEDI_PUBL']", Editorial);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Editorial", true, file);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtPUB_ISBN']", ISBN);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("ISBN", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlPUB_CLAS']", TipoPublicacion);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Tipo Publicacion", true, file);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFEC_PUBL_txtFecha']", Fecha);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Fecha", true, file);
                                    //Guardar
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Registro Guardado", true, file);
                                    //Registro a editar
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgPubli_ctl03_LinkButton1']/i");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Registro a editar", true, file);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtTit_Publ']", Titulo + "Editado");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Titulo Editado", true, file);
                                    //Actualizar
                                    selenium.Click("//a[@id='btnActualizar']");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Registro Actualizado", true, file);
                                    //Eliminar Registro
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgPubli_ctl03_LinkButton2']/i");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Registro Eliminado", true, file);
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Reclutamiento.Trayectoria")
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
                                rows["RlUser"].ToString().Length != 0 && rows["RlUser"].ToString() != null &&
                                rows["RlPass"].ToString().Length != 0 && rows["RlPass"].ToString() != null
                                // Data Trayectoria RL/////////////////////

                                )
                            {
                                //LOGIN
                                string RlUser = rows["RlUser"].ToString();
                                string RlPass = rows["RlPass"].ToString();
                                // Data Trayectoria RL/////////////////////
                                string Institucion = rows["Institucion"].ToString();
                                string Autoria = rows["Autoria"].ToString();
                                string Proyecto = rows["Proyecto"].ToString();
                                string FechaInicio = rows["FechaInicio"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Url = rows["Url"].ToString();

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
                                    string user = "";
                                    string database = "";
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestAuto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    if (Url.ToLower() == "http://dwtfskscm/ReclutamientoTestoraAuto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                    }

                                    if (Url.ToLower() == "http://dwtfsk:8088/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                    }
                                    if (Url.ToLower() == "http://dwtfsk:8087/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();
                                    //BD
                                    //PARAMETRIZACION PREVIA
                                    if (database == "SQL")
                                    {
                                        string deletedocu = $"DELETE FROM RL_DTRIN WHERE COD_EMPR = '9'";
                                        db.UpdateDeleteInsert(deletedocu, database, user);

                                        string delete = $"DELETE FROM Rl_Trinv WHERE COD_EMPL = '845267621'";
                                        db.UpdateDeleteInsert(delete, database, user);

                                    }
                                    else if (database == "ORA")
                                    {
                                        string deletedocu = $"DELETE FROM RL_DTRIN WHERE COD_EMPR = '421'";
                                        db.UpdateDeleteInsert(deletedocu, database, user);

                                        string delete = $"DELETE FROM Rl_Trinv WHERE COD_EMPL = '1011121314'";
                                        db.UpdateDeleteInsert(delete, database, user);

                                    }

                                    //login
                                    selenium.LoginApps(app, RlUser, RlPass, Url, file);
                                    selenium.Screenshot("Login", true, file);
                                    Thread.Sleep(3000);
                                    //Trayectoria Investigativa
                                    selenium.Click("//a[contains(.,'Trayectoria Investigativa')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Trayectoria Investigativa", true, file);
                                    //Nuevo
                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nueva Trayectoria Investigativa", true, file);
                                    //---------------------------------------------------------Pestaña Trayectoria Investigativa-------------------------------------------------------------------------------------------
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Trayectoria a registrar", true, file);
                                    //Diligenciar
                                    selenium.Click("//a[contains(text(),'Trayectoria investigativa')]");
                                    Thread.Sleep(3000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnNuevoDtrinI']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nueva Trayectoria Investigativa", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomInstDtrinI']", Institucion);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtAutOriaI']", Autoria);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomProyI']", Proyecto);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInicI_txtFecha']", FechaInicio);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecFinaI_txtFecha']", FechaFinal);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Datos Trayectoria Investigativa", true, file);
                                    //Guardar
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGuardarDtrinI']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Guardado Trayectoria Investigativa", true, file);
                                    //Editar registro
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrBiDtrinI_ctl02_LinkButton21']/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Edicion Trayectoria Investigativa", true, file);
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomInstDtrinI']", Institucion + " Edicion");
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(3000);
                                    //Actualizar
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnActualizarDtrinI']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Editado Trayectoria Investigativa", true, file);
                                    Thread.Sleep(3000);
                                    //Eliminar Registro
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrBiDtrinI_ctl02_LinkButton22']/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Eliminado Trayectoria Investigativa", true, file);
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


    }
}
