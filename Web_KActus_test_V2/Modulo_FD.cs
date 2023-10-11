using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using OpenQA.Selenium.Chrome;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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
    /// <summary>
    /// Descripción resumida de SelfServices
    /// </summary>

    [TestClass]
    public class Modulo_FD : FuncionesVitales
    {

        string Modulo = "Modulo_FD";
        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();

        public Modulo_FD()
        {

        }

        [TestMethod]
        public void FD_EnvíoCorreoMisCursos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_EnvíoCorreoMisCursos")
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
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["EmailRecepcion"].ToString().Length != 0 && rows["EmailRecepcion"].ToString() != null &&
                                rows["PassEmailRecepcion"].ToString().Length != 0 && rows["PassEmailRecepcion"].ToString() != null &&
                                rows["Remite"].ToString().Length != 0 && rows["Remite"].ToString() != null &&
                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["ConfAsis"].ToString().Length != 0 && rows["ConfAsis"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["Parametro"].ToString().Length != 0 && rows["Parametro"].ToString() != null &&
                                rows["Secuencial"].ToString().Length != 0 && rows["Secuencial"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string EmailRecepcion = rows["EmailRecepcion"].ToString();
                                string PassEmailRecepcion = rows["PassEmailRecepcion"].ToString();
                                string Remite = rows["Remite"].ToString();
                                string Descripcion = rows["Descripcion"].ToString();
                                string CodPagina = rows["CodPagina"].ToString();
                                string ConfAsis = rows["ConfAsis"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string Parametro = rows["Parametro"].ToString();
                                string Secuencial = rows["Secuencial"].ToString();
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

                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {
                                        string actualizar = $"update FD_DPLIN set CON_FIRM='N',ASI_STIO='N', DPL_APTO='N' where IDE_NTIF='51880161' and RMT_PARA='73'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                    }
                                    else
                                    {

                                        string actualizar = $"update FD_DPLIN set CON_FIRM='N',ASI_STIO='N', DPL_APTO='N' where IDE_NTIF='15' and RMT_PARA='178'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                    }

                                    //Login
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    //Ingresar a confirmar Asistencia curso
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Confirmación a Cursos')]");
                                    selenium.Screenshot("Mis cursos", true, file);

                                   
                                        if (selenium.ExistControl("//select[contains(@id,'ddlConfCurso')]"))
                                        {
                                            //CONFIRMACION CURSO
                                            Thread.Sleep(200);
                                            selenium.SelectElementByName("//select[contains(@id,'ddlConfCurso')]", ConfAsis);
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Seleccionar confirmación curso", true, file);
                                            //ADICIONAR
                                            Thread.Sleep(1000);
                                            selenium.Click("//*[@id='ctl00_ContenidoPagina_Adicionar']");
                                            Thread.Sleep(2000);
                                            if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                            {
                                                selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                                Thread.Sleep(2000);
                                            }

                                            selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                            Thread.Sleep(500);
                                            selenium.Click("//a[contains(.,'Confirmación a Cursos')]");
                                            selenium.Screenshot("Mis cursos", true, file);
                                            selenium.Screenshot("Confirmada Asistencia", true, file);

                                        }
                                        else
                                        {
                                            selenium.Screenshot("No hay cursos por confirmar", true, file);
                                            Assert.Fail("No hay cursos por confirmar");
                                        }

                                   
                                    Thread.Sleep(5000);
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
        public void FD_EnvíoCorreoSolicitudFormación()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_EnvíoCorreoSolicitudFormación")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["EmailRecepcion"].ToString().Length != 0 && rows["EmailRecepcion"].ToString() != null &&
                                rows["PassEmailRecepcion"].ToString().Length != 0 && rows["PassEmailRecepcion"].ToString() != null &&
                                rows["Remite"].ToString().Length != 0 && rows["Remite"].ToString() != null &&
                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["Registro"].ToString().Length != 0 && rows["Registro"].ToString() != null &&
                                rows["Curso"].ToString().Length != 0 && rows["Curso"].ToString() != null &&
                                rows["Perspectiva"].ToString().Length != 0 && rows["Perspectiva"].ToString() != null &&
                                rows["Justificacion"].ToString().Length != 0 && rows["Justificacion"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["TipoDocumAdjunto"].ToString().Length != 0 && rows["TipoDocumAdjunto"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string CodPagina = rows["CodPagina"].ToString();
                                string EmailRecepcion = rows["EmailRecepcion"].ToString();
                                string PassEmailRecepcion = rows["PassEmailRecepcion"].ToString();
                                string Remite = rows["Remite"].ToString();
                                string Descripcion = rows["Descripcion"].ToString();
                                string Registro = rows["Registro"].ToString();
                                string Curso = rows["Curso"].ToString();
                                string Perspectiva = rows["Perspectiva"].ToString();
                                string Justificacion = rows["Justificacion"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string TipoDocumAdjunto = rows["TipoDocumAdjunto"].ToString();
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

                                    // ELIMINAR REGISTROS PREVIOS
                                    //eliminar registros previos

                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarNecesidad = $"Delete from FD_NEFOR where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarNecesidad, database, user);

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A MIS CURSOS
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(500);
                                    if (database == "SQL")
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                    }
                                    else
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                    }
                                    Thread.Sleep(1000);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    selenium.Screenshot("Crear registro", true, file);
                                    Thread.Sleep(2000);
                                    //REGISTRO
                                    selenium.Scroll("//select[contains(@id,'ddlNomRegi')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomRegi')]", Registro);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro", true, file);
                                    //CURSOS
                                    selenium.Scroll("//select[contains(@id,'ddlNomCurs')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomCurs')]", Curso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Curso", true, file);
                                    //PERSPECTIVA
                                    selenium.Scroll("//select[contains(@id,'ddlCodPers')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodPers')]", Perspectiva);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Perspectiva", true, file);
                                    //JUSTIFICACION
                                    selenium.Scroll("//textarea[contains(@id,'txtJusSoli_txtTexto')]");
                                    selenium.SendKeys("//textarea[contains(@id,'txtJusSoli_txtTexto')]", Justificacion);
                                    Thread.Sleep(2000);
                                    //OBSERVACION
                                    selenium.Scroll("//textarea[contains(@id,'txtObsErva_txtTexto')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//textarea[contains(@id,'txtObsErva_txtTexto')]");
                                    Thread.Sleep(2500);
                                    selenium.SendKeys("//textarea[contains(@id,'txtObsErva_txtTexto')]", Observacion);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Datos Necesidad de formación", true, file);
                                    Thread.Sleep(2000);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.AcceptAlert();
                                    selenium.Screenshot("Necesidad de formación Enviada", true, file);
                                    Thread.Sleep(3000);
                                    //Necesidades de formacion
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(2500);
                                    if (database == "SQL")
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                        selenium.Screenshot("Necesidad de Formacion Registrada", true, file);
                                    }
                                    else
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                        selenium.Screenshot("Necesidad de formación Registrada", true, file);
                                    }
                                    Thread.Sleep(1000);
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
        public void FD_ReporteAsistenciaCursos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_ReporteAsistenciaCursos")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["EmailRecepcion"].ToString().Length != 0 && rows["EmailRecepcion"].ToString() != null &&
                                rows["PassEmailRecepcion"].ToString().Length != 0 && rows["PassEmailRecepcion"].ToString() != null &&
                                rows["Remite"].ToString().Length != 0 && rows["Remite"].ToString() != null &&
                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["CodEmple"].ToString().Length != 0 && rows["CodEmple"].ToString() != null &&
                                rows["CodEmpre"].ToString().Length != 0 && rows["CodEmpre"].ToString() != null &&
                                rows["Parametro"].ToString().Length != 0 && rows["Parametro"].ToString() != null &&
                                rows["Secuencial"].ToString().Length != 0 && rows["Secuencial"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string CodPagina = rows["CodPagina"].ToString();
                                string EmailRecepcion = rows["EmailRecepcion"].ToString();
                                string PassEmailRecepcion = rows["PassEmailRecepcion"].ToString();
                                string Remite = rows["Remite"].ToString();
                                string Descripcion = rows["Descripcion"].ToString();
                                string CodEmple = rows["CodEmple"].ToString();
                                string CodEmpre = rows["CodEmpre"].ToString();
                                string Parametro = rows["Parametro"].ToString();
                                string Secuencial = rows["Secuencial"].ToString();
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
                                    string username = Environment.UserName;
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    string ActualizarRegistro = $"UPDATE FD_DPLIN SET ASI_STIO = 'N' WHERE RMT_PARA = {Parametro} AND RMT_PLCU = {Secuencial} AND COD_EMPR = {CodEmpre} AND IDE_NTIF IN{CodEmple}";
                                    db.UpdateDeleteInsert(ActualizarRegistro, database, user);

                                    File.Delete("C:/Users/" + username + "/Downloads/AsistenciaCurso.pdf");

                                    //INGRESAR MIS CURSOS
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(500);
                                    if (database == "SQL")
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[7]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[7]/a");
                                    }
                                    else
                                    {
                                        selenium.Scroll("//*[@id='MenuContex']/div[2]/div[1]/ul/li[4]/ul/li[11]/a");
                                        selenium.Click("//*[@id='MenuContex']/div[2]/div[1]/ul/li[4]/ul/li[11]/a");
                                    }

                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Reportes de Asistencia a Cursos", true, file);

                                    //DETALLE
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFdDplse_ctl03_lbSelDetalle']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle Seleccionado", true, file);

                                    //REPORTE
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnReporte']");
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnReporte']");
                                    Thread.Sleep(2000);
                                    Screenshot("Reporte Asistencia Generado", true, file);

                                    //VENTANA EMERGENTE
                                    String mainWin = selenium.MainWindow();
                                    String modalWin = selenium.PopupWindow();
                                    selenium.ChangeWindow(modalWin);
                                    Thread.Sleep(5000);
                                    Screenshot("Reporte Asistencia Generado", true, file);
                                    selenium.MaximizeWindow();
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Reporte Asistencia Cursos", true, file);
                                    Thread.Sleep(5000);
                                    //PDF IMPRIMIR
                                    selenium.Click("//*[@id='ctl00_btnImprimir']");
                                    Thread.Sleep(20000);
                                    Screenshot("IMPRIMIR PDF", true, file);
                                    //GUARDAR PDF
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }

                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{DOWN}");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);

                                    for (int i = 0; i < 6; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("AsistenciaCurso");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);

                                    //ABRIR PDF
                                    string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/" + username + "/Downloads/AsistenciaCurso.pdf");
                                    Process.Start(pdfPath);
                                    Thread.Sleep(6000);
                                    Screenshot("PDF ABIERTO", true, file);
                                    KillProcesos("Acrobat.exe");
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    Thread.Sleep(2000);
                                    selenium.ChangeWindow(mainWin);
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

        public void FD_FlujoAprobaciónNecesidadesFormaciónRolLider()
        {
            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FlujoAprobaciónNecesidadesFormaciónRolLider")
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
                                //Datos Necesidades Formación   
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["JefeUser"].ToString().Length != 0 && rows["JefeUser"].ToString() != null &&
                                rows["JefePass"].ToString().Length != 0 && rows["JefePass"].ToString() != null &&
                                rows["Registro"].ToString().Length != 0 && rows["Registro"].ToString() != null &&
                                rows["Curso"].ToString().Length != 0 && rows["Curso"].ToString() != null &&
                                rows["Perspectiva"].ToString().Length != 0 && rows["Perspectiva"].ToString() != null &&
                                rows["Justificacion"].ToString().Length != 0 && rows["Justificacion"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CentroCosto"].ToString().Length != 0 && rows["CentroCosto"].ToString() != null &&
                                rows["CodCurso"].ToString().Length != 0 && rows["CodCurso"].ToString() != null &&
                                rows["TipoDocumentAdjunto"].ToString().Length != 0 && rows["TipoDocumentAdjunto"].ToString() != null
                                //rows["CodNiv2"].ToString().Length != 0 && rows["CodNiv2"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string CodPagina = rows["CodPagina"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();
                                string Registro = rows["Registro"].ToString();
                                string Curso = rows["Curso"].ToString();
                                string Perspectiva = rows["Perspectiva"].ToString();
                                string Justificacion = rows["Justificacion"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CentroCosto = rows["CentroCosto"].ToString();
                                string CodCurso = rows["CodCurso"].ToString();
                                //string CodNiv2 = rows["CodNiv2"].ToString();
                                //string EstadoFormacion = rows["EstadoFormacion"].ToString();
                                string TipoDocumentAdjunto = rows["TipoDocumentAdjunto"].ToString();
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
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarNecesidad = $"Delete from FD_NEFOR where COD_EMPL ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarNecesidad, database, user);

                                    string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{JefeUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud1, database, user);


                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    //INGRESO A MIS CURSOS
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(500);

                                    if (database == "SQL")
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                    }
                                    else
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                    }

                                    Thread.Sleep(1000);

                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    selenium.Screenshot("Crear registro", true, file);
                                    //REGISTRO
                                    selenium.Scroll("//select[contains(@id,'ddlNomRegi')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomRegi')]", Registro);
                                    Thread.Sleep(2200);
                                    selenium.Screenshot("Registro", true, file);
                                    //CURSO
                                    selenium.Scroll("//select[contains(@id,'ddlNomCurs')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomCurs')]", Curso);
                                    Thread.Sleep(2200);
                                    selenium.Screenshot("Curso", true, file);
                                    //PERSPECTIVA
                                    selenium.Scroll("//select[contains(@id,'ddlCodPers')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodPers')]", Perspectiva);
                                    Thread.Sleep(2200);
                                    selenium.Screenshot("Perspectiva", true, file);
                                    //JUSTIFICACION
                                    selenium.Scroll("//textarea[contains(@id,'txtJusSoli_txtTexto')]");
                                    selenium.SendKeys("//textarea[contains(@id,'txtJusSoli_txtTexto')]", Justificacion);
                                    Thread.Sleep(2200);
                                    selenium.Screenshot("Justificación", true, file);
                                    //OBSERVACIONES 
                                    Thread.Sleep(200);
                                    selenium.Scroll("//textarea[contains(@id,'txtObsErva_txtTexto')]");
                                    selenium.SendKeys("//textarea[contains(@id,'txtObsErva_txtTexto')]", Observacion);
                                    selenium.Screenshot("Datos diligenciados en Formulario", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    //INGRESO A MIS CURSOS
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(5000);

                                    if (database == "SQL")
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                    }
                                    else
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                    }

                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Necesidad de Formación Enviada", true, file);
                                    Thread.Sleep(4000);
                                    selenium.Close();

                                    if (database == "SQL")
                                    {
                                        //APERTURA LIDER PARA APROBACION DE LA FORMACION
                                        Thread.Sleep(2000);
                                        selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                        Thread.Sleep(500);
                                        selenium.Click("//span[contains(@id,'pLider')]");
                                        selenium.Screenshot("Rol Lider", true, file);
                                        Thread.Sleep(500);
                                        selenium.Click("//a[contains(.,'FORMACIÓN Y DESARROLLO')]");
                                        selenium.Screenshot("Formación y Desarrollo", true, file);
                                        Thread.Sleep(200);
                                        selenium.Scroll("//a[contains(.,'Aprobación de N. Formación')]");
                                        Thread.Sleep(500);
                                        selenium.Click("//a[contains(.,'Aprobación de N. Formación')]");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Aprobación de Necesidades de Formación", true, file);
                                        Thread.Sleep(500);

                                        if (selenium.ExistControl("//a[contains(@id,'ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1')]"))
                                        {
                                            selenium.Scroll("//a[contains(@id,'ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1')]");
                                            selenium.Click("//a[contains(@id,'ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1')]");
                                            Thread.Sleep(500);
                                            selenium.Screenshot("Detalle Aprobación de Necesidades de Formación", true, file);
                                            selenium.Click("//div[@id='ctl00_pBotones']/div");
                                            selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_btnAprobar')]");
                                            Thread.Sleep(2000);
                                            selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnAprobar')]");
                                            Thread.Sleep(2000);
                                            Screenshot("Alerta Aprobación 1", true, file);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(2000);
                                            Screenshot("Alerta Aprobación 2", true, file);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(1000);
                                            selenium.Screenshot("Acción Exitosa", true, file);
                                        }
                                        else
                                        {
                                            Assert.Fail("ERROR: NO HAY NECESIDADES PENDIENTES POR APROBAR DE MIS COLABORADORES");
                                        }

                                        selenium.Close();

                                    }
                                    else
                                    {
                                        //APERTURA LIDER PARA APROBACION DE LA FORMACION
                                        Thread.Sleep(2000);
                                        selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                        Thread.Sleep(500);
                                        selenium.Click("//span[contains(@id,'pLider')]");
                                        selenium.Screenshot("Rol Lider", true, file);
                                        Thread.Sleep(500);
                                        selenium.Click("//a[contains(.,'FORMACIÓN Y DESARROLLO')]");
                                        selenium.Screenshot("Formación y Desarrollo", true, file);
                                        Thread.Sleep(200);
                                        selenium.Scroll("//a[contains(.,'Aprobación de N. Formación')]");
                                        Thread.Sleep(500);
                                        selenium.Click("//a[contains(.,'Aprobación de N. Formación')]");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Aprobación de Necesidades de Formación", true, file);

                                        Thread.Sleep(500);

                                        if (selenium.ExistControl("//a[contains(@id,'ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1')]"))
                                        {
                                            selenium.Scroll("//a[contains(@id,'ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1')]");
                                            selenium.Click("//a[contains(@id,'ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1')]");
                                            Thread.Sleep(500);
                                            selenium.Screenshot("Detalle Aprobación de Necesidades de Formación", true, file);
                                            selenium.Click("//div[@id='ctl00_pBotones']/div");
                                            selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_btnAprobar')]");
                                            Thread.Sleep(2000);
                                            selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnAprobar')]");
                                            Thread.Sleep(2000);
                                            Screenshot("Alerta Aprobación 1", true, file);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(2000);
                                            Screenshot("Alerta Aprobación 2", true, file);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(1000);
                                            selenium.Screenshot("Acción Exitosa", true, file);
                                        }
                                        else
                                        {
                                            Assert.Fail("ERROR: NO HAY NECESIDADES PENDIENTES POR APROBAR DE MIS COLABORADORES");
                                        }

                                        selenium.Close();
                                    }

                                    // Abrir para verificar que la Necesidad de Formación fue aprobada 
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    selenium.Screenshot("Mis Cursos", true, file);
                                    Thread.Sleep(500);

                                    if (database == "SQL")
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                    }
                                    else
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                    }

                                    selenium.Screenshot("Necesidades Formación", true, file);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Necesidades Formación Aprobadas", true, file);
                                    Thread.Sleep(3000);
                                    ////////////////////////////////
                                    Thread.Sleep(5000);
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

        public void FD_FlujoRechazoNecesidadesFormaciónRolLider()
        {
            List<string> errorsTest = new List<string>();
            List<string> errors = new List<string>();
            List<string> CargarDocumentos = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FlujoRechazoNecesidadesFormaciónRolLider")
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
                                //Datos Necesidades Formación Rechazo
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["JefeUser"].ToString().Length != 0 && rows["JefeUser"].ToString() != null &&
                                rows["JefePass"].ToString().Length != 0 && rows["JefePass"].ToString() != null &&
                                rows["Registro"].ToString().Length != 0 && rows["Registro"].ToString() != null &&
                                rows["Curso"].ToString().Length != 0 && rows["Curso"].ToString() != null &&
                                rows["Perspectiva"].ToString().Length != 0 && rows["Perspectiva"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CentroCosto"].ToString().Length != 0 && rows["CentroCosto"].ToString() != null &&
                                //rows["CodPersona"].ToString().Length != 0 && rows["CodPersona"].ToString() != null &&
                                //rows["EstadoFormacion"].ToString().Length != 0 && rows["EstadoFormacion"].ToString() != null &&
                                rows["TipoDocumentAdjunto"].ToString().Length != 0 && rows["TipoDocumentAdjunto"].ToString() != null &&
                                rows["CodCurso"].ToString().Length != 0 && rows["CodCurso"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string CodPagina = rows["CodPagina"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();
                                string Registro = rows["Registro"].ToString();
                                string Curso = rows["Curso"].ToString();
                                string Perspectiva = rows["Perspectiva"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CentroCosto = rows["CentroCosto"].ToString();
                                //string CodPersona = rows["CodPersona"].ToString();
                                //string EstadoFormacion = rows["EstadoFormacion"].ToString();
                                string CodCurso = rows["CodCurso"].ToString();
                                string TipoDocumentAdjunto = rows["TipoDocumentAdjunto"].ToString();
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

                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarNecesidad = $"Delete from FD_NEFOR where COD_EMPL ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarNecesidad, database, user);

                                    string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{JefeUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    //INGRESO A MIS CURSOS
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(500);

                                    if (database == "SQL")
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                        Thread.Sleep(1500);
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                    }
                                    else
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                        Thread.Sleep(1500);
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                    }
                                    Thread.Sleep(1000);

                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    selenium.Screenshot("Crear registro", true, file);
                                    Thread.Sleep(2000);
                                    //REGISTRO
                                    selenium.Scroll("//select[contains(@id,'ddlNomRegi')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomRegi')]", Registro);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro", true, file);
                                    //CURSO
                                    selenium.Scroll("//select[contains(@id,'ddlNomCurs')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomCurs')]", Curso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Curso", true, file);
                                    //PERSPECTIVA
                                    selenium.Scroll("//select[contains(@id,'ddlCodPers')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodPers')]", Perspectiva);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Perspectiva", true, file);
                                    //OBSERVACIONES
                                    selenium.Scroll("//textarea[contains(@id,'txtObsErva_txtTexto')]");
                                    selenium.Click("//textarea[contains(@id,'txtObsErva_txtTexto')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[contains(@id,'txtObsErva_txtTexto')]", Observacion);
                                    selenium.Screenshot("Datos diligenciados en Formulario", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(5000);

                                    //INGRESO A MIS CURSOS
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(500);

                                    if (database == "SQL")
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                        Thread.Sleep(1500);
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                    }
                                    else
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                        Thread.Sleep(1500);
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                    }
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Necesidad de Formación Enviada", true, file);
                                    Thread.Sleep(5000);
                                    selenium.Close();

                                    if (database == "SQL")
                                    {
                                        //ABRIR LIDER PARA RECHAZAR LA NECESIDAD DE FORMACION
                                        Thread.Sleep(2000);
                                        selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                        Thread.Sleep(1500);
                                        selenium.Click("//span[contains(@id,'pLider')]");
                                        selenium.Screenshot("Rol Lider", true, file);
                                        Thread.Sleep(1500);
                                        selenium.Click("//a[contains(.,'FORMACIÓN Y DESARROLLO')]");
                                        selenium.Screenshot("Formación y Desarrollo", true, file);
                                        Thread.Sleep(1200);
                                        selenium.Scroll("//a[contains(.,'Aprobación de N. Formación')]");
                                        Thread.Sleep(1500);
                                        selenium.Click("//a[contains(.,'Aprobación de N. Formación')]");
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Aprobación de Necesidades de Formación", true, file);
                                        Thread.Sleep(1500);

                                        if (selenium.ExistControl("//a[contains(@id,'ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1')]"))
                                        {
                                            selenium.Scroll("//a[contains(@id,'ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1')]");
                                            selenium.Click("//a[contains(@id,'ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1')]");
                                            Thread.Sleep(1500);
                                            selenium.Click("//div[@id='ctl00_pBotones']/div");
                                            selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_btnRechasa')]");
                                            Thread.Sleep(2000);
                                            selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnRechasa')]");
                                            Screenshot("Alerta Rechazo", true, file);
                                            Thread.Sleep(5000);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Rechazo Exitoso", true, file);

                                        }
                                        else
                                        {
                                            Assert.Fail("ERROR: NO HAY NECESIDADES PENDIENTES POR APROBAR DE MIS COLABORADORES");
                                        }

                                        Thread.Sleep(3000);
                                        selenium.Close();
                                    }
                                    else
                                    {
                                        //ABRIR LIDER PARA RECHAZAR LA NECESIDAD DE FORMACION
                                        Thread.Sleep(2000);
                                        selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                        Thread.Sleep(1500);
                                        selenium.Click("//span[contains(@id,'pLider')]");
                                        selenium.Screenshot("Rol Lider", true, file);
                                        Thread.Sleep(1500);
                                        selenium.Click("//a[contains(.,'FORMACIÓN Y DESARROLLO')]");
                                        selenium.Screenshot("Formación y Desarrollo", true, file);
                                        Thread.Sleep(1200);
                                        selenium.Scroll("//a[contains(.,'Aprobación de N. Formación')]");
                                        Thread.Sleep(1500);
                                        selenium.Click("//a[contains(.,'Aprobación de N. Formación')]");
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Aprobación de Necesidades de Formación", true, file);

                                        Thread.Sleep(1500);

                                        if (selenium.ExistControl("//a[contains(@id,'ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1')]"))
                                        {
                                            selenium.Scroll("//a[contains(@id,'ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1')]");
                                            selenium.Click("//a[contains(@id,'ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1')]");
                                            Thread.Sleep(1500);
                                            selenium.Click("//div[@id='ctl00_pBotones']/div");
                                            selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_btnRechasa')]");
                                            Thread.Sleep(2000);
                                            selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnRechasa')]");
                                            Thread.Sleep(5000);
                                            Screenshot("Alerta Rechazo", true, file);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Rechazo Exitoso", true, file);
                                        }
                                        else
                                        {
                                            Assert.Fail("ERROR: NO HAY NECESIDADES PENDIENTES POR APROBAR DE MIS COLABORADORES");
                                        }
                                        selenium.Close();


                                    }

                                    // Abrir para verificar que la Necesidad de Formación fue Rechazada 
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    selenium.Screenshot("Mis Cursos", true, file);
                                    Thread.Sleep(1500);

                                    if (database == "SQL")
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[9]/a");
                                    }
                                    else
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[8]/a");
                                    }
                                    selenium.Screenshot("Necesidades Formación", true, file);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Necesidades Formación Aprobadas", true, file);
                                    Thread.Sleep(3000);
                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(5000);
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
        public void FD_FormaciónDesarrolloConfirmaciónCursos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloConfirmaciónCursos")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba
                                rows["Confirmacion"].ToString().Length != 0 && rows["Confirmacion"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Confirmacion = rows["Confirmacion"].ToString();
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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }


                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {
                                        string actualizar = $"update FD_DPLIN set CON_FIRM='N',ASI_STIO='N', DPL_APTO='N' where IDE_NTIF='51880161' and RMT_PARA='73'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                    }
                                    else
                                    {

                                        string actualizar = $"update FD_DPLIN set CON_FIRM='N',ASI_STIO='N', DPL_APTO='N' where IDE_NTIF='123' and RMT_PARA='178'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
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

                                    //INGRESO A MIS CURSOS
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Confirmación a Cursos')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A CONFIRMACION CURSOS", true, file);

                                    if (selenium.ExistControl("//select[@id='ctl00_ContenidoPagina_dtgFdCoCur_ctl02_ddlConfCurso']"))
                                        {
                                            //CONFIRMAR ASISTENCIA
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_dtgFdCoCur_ctl02_ddlConfCurso']", Confirmacion);
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Confirmar asistencia", true, file);
                                            //GUARDAR
                                            selenium.Click("//*[@id='ctl00_ContenidoPagina_Adicionar']");
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Confirmacion exitosa", true, file);

                                            Thread.Sleep(5000);
                                            selenium.Close();

                                    }
                                        
                                    else
                                    {
                                        Assert.Fail("No hay cursos por confirmar");
                                    }
                                    
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
        public void FD_FormaciónDesarrolloCertificadoCursos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloCertificadoCursos" +
                    "")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba
                                rows["Certificado"].ToString().Length != 0 && rows["Certificado"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Certificado = rows["Certificado"].ToString();
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


                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {

                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='E' WHERE RMT_PARA='73'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
                                        string actualizar = $"update FD_DPLIN set ASI_STIO='S' where IDE_NTIF='51880161' and RMT_PARA='73'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                    }
                                    else
                                    {

                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='E' WHERE RMT_PARA='178'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
                                        string actualizar = $"update FD_DPLIN set ASI_STIO='S' where IDE_NTIF='123' and RMT_PARA='178'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                    }

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    string username = Environment.UserName;
                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               
                                    //BORRADO PDF
                                    File.Delete("C:/Users/" + username + "/Downloads/CertificadoCursos.pdf");
                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A MIS CURSOS
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Certificados Cursos')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Certificado Cursos", true, file);

                                    //SELECCIONAR DETALLE
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdPlcur_ctl03_lbSelDetalle']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Certificado a generar", true, file);
                                    //SELECCIONAR TIPO CERTIFICADO
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCerCurs']", Certificado);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Certificado a generar", true, file);
                                    //GENERAR
                                    selenium.Click("//a[@id='ctl00_btnGenerar']");
                                    Thread.Sleep(5000);
                                    //VENTANA EMERGENTE
                                    String mainWin = selenium.MainWindow();
                                    String modalWin = selenium.PopupWindow();
                                    Thread.Sleep(3000);
                                    selenium.ChangeWindow(modalWin);
                                    Thread.Sleep(5000);
                                    Screenshot("Certificado Generado", true, file);
                                    selenium.MaximizeWindow();
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Certificado", true, file);
                                    Thread.Sleep(5000);
                                    //PDF IMPRIMIR
                                    selenium.Click("//a[@id='imprimir']");
                                    Thread.Sleep(20000);
                                    Screenshot("IMPRIMIR PDF", true, file);
                                    //GUARDAR PDF
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }

                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{DOWN}");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);

                                    for (int i = 0; i < 6; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("CertificadoCursos");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);

                                    //ABRIR PDF
                                    string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/" + username + "/Downloads/CertificadoCursos.pdf");
                                    Process.Start(pdfPath);
                                    Thread.Sleep(6000);
                                    Screenshot("PDF ABIERTO", true, file);
                                    Thread.Sleep(5000);
                                    selenium.Close();
                                    selenium.ChangeWindow(mainWin);
                                    Thread.Sleep(5000);
                                    selenium.Close();
                                    

                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {

                                        string actualizar = $"update FD_PLCUR SET EST_CURS='P' WHERE RMT_PARA='73'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                    }
                                    else
                                    {

                                        string actualizar = $"update FD_PLCUR SET EST_CURS='P' WHERE RMT_PARA='178'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                    }
                                  
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
                                    Thread.Sleep(5000);
                                    KillProcesos("Acrobat.exe");
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
        public void FD_FormaciónDesarrolloInscripción()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloInscripción")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                //Datos Prueba



                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
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

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {
                                        string ActualizarCheck = $"Update FD_PARAM set INS_JEFE ='N' Where COD_CURS= 5";
                                        db.UpdateDeleteInsert(ActualizarCheck, database, user);
                                        string BorrarREgistro = $"delete from fd_dplin where rmt_para = 73 and ide_ntif = 193454 and cod_empr = 421";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
                                        string BorrarRegistro = $"update FD_PLCUR SET EST_CURS='P' WHERE RMT_PARA='73'";
                                        db.UpdateDeleteInsert(BorrarRegistro, database, user);
                                    }
                                    else
                                    {
                                        string ActualizarCheck = $"Update FD_PARAM set INS_JEFE ='N' Where COD_CURS= 5";
                                        db.UpdateDeleteInsert(ActualizarCheck, database, user);
                                        string BorrarREgistro = $"delete from fd_dplin where rmt_para = 178 and ide_ntif = 4 and cod_empr = 9";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
                                        string BorrarRegistro = $"update FD_PLCUR SET EST_CURS='P' WHERE RMT_PARA='178'";
                                        db.UpdateDeleteInsert(BorrarRegistro, database, user);
                                    }

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A MIS CURSOS/INSCRIPCION
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Inscripción a Cursos')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A INSCRIPCION CURSOS", true, file);

                                    //DETALLE
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFdPlcur_ctl03_LinkButton1']/i[1]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle", true, file);

                                    //INSCRIBIRSE
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_btnInscribirse']");
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_btnInscribirse']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Inscrito", true, file);

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(5000);
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
        public void FD_FormaciónDesarrolloInscripciónWebJefe()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloInscripciónWebJefe")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                //Datos Prueba



                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
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

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {
                                        string ActualizarCheck = $"Update FD_PARAM set INS_JEFE ='S' Where COD_CURS= 5 and cod_prog= 8";
                                        db.UpdateDeleteInsert(ActualizarCheck, database, user);
                                        string BorrarREgistro = $"delete from fd_dplin where rmt_para = 73 and ide_ntif = 193454 and cod_empr = 421";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
                                        string BorrarRegistro = $"update FD_PLCUR SET EST_CURS='P' WHERE RMT_PARA='73'";
                                        db.UpdateDeleteInsert(BorrarRegistro, database, user);
                                    }
                                    else
                                    {
                                        string ActualizarCheck = $"Update FD_PARAM set INS_JEFE ='S' Where COD_CURS= 5 and cod_prog= 8";
                                        db.UpdateDeleteInsert(ActualizarCheck, database, user);
                                        string BorrarREgistro = $"delete from fd_dplin where rmt_para = 178 and ide_ntif = 4 and cod_empr = 9";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
                                        string BorrarRegistro = $"update FD_PLCUR SET EST_CURS='P' WHERE RMT_PARA='178'";
                                        db.UpdateDeleteInsert(BorrarRegistro, database, user);
                                    }

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A ROL LIDER
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                        Thread.Sleep(4000);
                                        selenium.Screenshot("ROL LIDER", true, file);
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'Lider')]");
                                        Thread.Sleep(4000);
                                        selenium.Screenshot("LIDER", true, file);
                                    }
                                    //INGRESO A FORMACION Y DESARROLLO/INSCRIPCION
                                    selenium.Click("//a[contains(.,'FORMACIÓN Y DESARROLLO')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Inscripcion a Cursos')]");
                                    selenium.Click("//a[contains(.,'Inscripcion a Cursos')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A INSCRIPCION CURSOS", true, file);

                                    //DETALLE
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdPlcur_ctl03_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    //INSCRIBIRSE PARTICIPANTES
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_LkbInscribir']");
                                    selenium.Screenshot("INSCRIBIR POSTULADOS ", true, file);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_LkbInscribir']");
                                    Thread.Sleep(2000);
                                    //INSCRIBIR
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_DdlBusqueda']");
                                    selenium.Screenshot("INSCRIPCION POSTULADO", true, file);
                                    Thread.Sleep(3000);

                                    if (database == "ORA")
                                    {
                                        //BUSQUEDA
                                        selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_DdlBusqueda']", "Identificación");
                                        Thread.Sleep(2000);
                                        selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_TextFiltro']", "193454");
                                        Thread.Sleep(2000);
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_ImageButton1']");
                                        Thread.Sleep(2000);
                                        
                                        selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgFdDplin_ctl03_btnInscribirse']/i[1]");
                                        selenium.Screenshot("POSTULADO", true, file);
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFdDplin_ctl03_btnInscribirse']/i[1]");
                                        Thread.Sleep(5000);
                                        selenium.Screenshot("INSCRIPCION EXITOSA", true, file);
                                    }
                                    else
                                    {
                                        //BUSQUEDA
                                        selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_DdlBusqueda']", "Identificación");
                                        Thread.Sleep(2000);
                                        selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_TextFiltro']", "4");
                                        Thread.Sleep(2000);
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_ImageButton1']");
                                        Thread.Sleep(2000);
                                        selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgFdDplin_ctl03_btnInscribirse']/i[1]");
                                        selenium.Screenshot("POSTULADO", true, file);
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFdDplin_ctl03_btnInscribirse']/i[1]");
                                        Thread.Sleep(5000);
                                        selenium.Screenshot("INSCRIPCION EXITOSA", true, file);
                                    }

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(5000);
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
        public void FD_FormaciónDesarrolloSolicitudesFormación()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloSolicitudesFormación")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&

                                //Datos Prueba
                                rows["Registro"].ToString().Length != 0 && rows["Registro"].ToString() != null &&
                                rows["Requerimiento"].ToString().Length != 0 && rows["Requerimiento"].ToString() != null &&
                                rows["Especificacion"].ToString().Length != 0 && rows["Especificacion"].ToString() != null &&
                                rows["Curso"].ToString().Length != 0 && rows["Curso"].ToString() != null



                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Registro = rows["Registro"].ToString();
                                string Requerimiento = rows["Requerimiento"].ToString();
                                string Especificacion = rows["Especificacion"].ToString();
                                string Curso = rows["Curso"].ToString();
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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
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

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //PARAMETRIZACION PREVIA


                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {

                                        string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarSolicitud, database, user);
                                        string eliminarNecesidad = $"Delete from FD_NEFOR where COD_EMPL ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarNecesidad, database, user);
                                        string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='41416618'";
                                        db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    }
                                    else
                                    {

                                        string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarSolicitud, database, user);
                                        string eliminarNecesidad = $"Delete from FD_NEFOR where COD_EMPL ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarNecesidad, database, user);
                                        string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='124'";
                                        db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    }
                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);


                                    //INGRESO A MIS CURSOS/NECESIDADES FORMACION
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Mis necesidades de Formación')]");
                                    selenium.Click("//a[contains(.,'Mis necesidades de Formación')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A NECESIDADES DE FORMACION", true, file);

                                    //NUEVO
                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    Thread.Sleep(3000);

                                    //REGISTRO
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomRegi']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomRegi']", Registro);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro", true, file);
                                    //REQUERIMIENTO
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomRequ']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomRequ']", Requerimiento);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Requerimiento", true, file);
                                    //ESPECIFICACION
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomEspe']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomEspe']", Especificacion);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Especificacion", true, file);
                                    //CURSOS
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomCurs']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomCurs']", Curso);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Curso", true, file);
                                    //APLICAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(3000);

                                    //ACEPTAR ALERTA
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);

                                    //VERIFICAR REGISTRO
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Mis necesidades de Formación')]");
                                    selenium.Click("//a[contains(.,'Mis necesidades de Formación')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("REGISTRO EXITOSO", true, file);

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(5000);
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
        public void FD_FormaciónDesarrolloSolicitudesFormaciónJefe()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloSolicitudesFormaciónJefe")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&

                                //Datos Prueba
                                rows["Registro"].ToString().Length != 0 && rows["Registro"].ToString() != null &&
                                rows["Requerimiento"].ToString().Length != 0 && rows["Requerimiento"].ToString() != null &&
                                rows["Especificacion"].ToString().Length != 0 && rows["Especificacion"].ToString() != null &&
                                rows["Curso"].ToString().Length != 0 && rows["Curso"].ToString() != null &&
                                rows["Perspectiva"].ToString().Length != 0 && rows["Perspectiva"].ToString() != null &&
                                rows["Objetivo"].ToString().Length != 0 && rows["Objetivo"].ToString() != null &&
                                rows["Intensidad"].ToString().Length != 0 && rows["Intensidad"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Valor"].ToString().Length != 0 && rows["Valor"].ToString() != null &&
                                rows["Ruta"].ToString().Length != 0 && rows["Ruta"].ToString() != null &&
                                rows["Empleado"].ToString().Length != 0 && rows["Empleado"].ToString() != null &&
                                rows["Entidad"].ToString().Length != 0 && rows["Entidad"].ToString() != null




                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Registro = rows["Registro"].ToString();
                                string Requerimiento = rows["Requerimiento"].ToString();
                                string Especificacion = rows["Especificacion"].ToString();
                                string Curso = rows["Curso"].ToString();
                                string Perspectiva = rows["Perspectiva"].ToString();
                                string Objetivo = rows["Objetivo"].ToString();
                                string Intensidad = rows["Intensidad"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Valor = rows["Valor"].ToString();
                                string Entidad = rows["Entidad"].ToString();
                                string Empleado = rows["Empleado"].ToString();
                                string Ruta = rows["Ruta"].ToString();

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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
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

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //PARAMETRIZACION PREVIA

                                    if (database == "ORA")
                                    {

                                        string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='51880161'";
                                        db.UpdateDeleteInsert(eliminarSolicitud, database, user);
                                        string eliminarNecesidad = $"Delete from FD_NEFOR where COD_EMPL ='51880161'";
                                        db.UpdateDeleteInsert(eliminarNecesidad, database, user);
                                        string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='41416618'";
                                        db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    }
                                    else
                                    {

                                        string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='123'";
                                        db.UpdateDeleteInsert(eliminarSolicitud, database, user);
                                        string eliminarNecesidad = $"Delete from FD_NEFOR where COD_EMPL ='123'";
                                        db.UpdateDeleteInsert(eliminarNecesidad, database, user);
                                        string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='124'";
                                        db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    }
                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A ROL LIDER
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                        Thread.Sleep(4000);
                                        selenium.Screenshot("ROL LIDER", true, file);
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'Lider')]");
                                        Thread.Sleep(4000);
                                        selenium.Screenshot("LIDER", true, file);
                                    }
                                    //INGRESO A FORMACION Y DESARROLLO/NECESIDADES
                                    selenium.Click("//a[contains(.,'FORMACIÓN Y DESARROLLO')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Solicitudes de Formacion')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A SOLICITUDES DE FORMACION", true, file);

                                    //NUEVO
                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    Thread.Sleep(3000);

                                    //REGISTRO
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomRegi']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomRegi']", Registro);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("REGISTRO", true, file);

                                    //REQUERIMIENTO
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomRequ']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomRequ']", Requerimiento);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Requerimiento", true, file);

                                    //ESPECIFICACION
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomEspe']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomEspe']", Especificacion);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Especificacion", true, file);

                                    //CURSOS
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomCurs']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomCurs']", Curso);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Curso", true, file);


                                    //PERSPECTIVA
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlCodPers']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodPers']", Perspectiva);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Perspectiva", true, file);

                                    //OBJETIVOS
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlCodObes']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodObes']", Objetivo);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Objetivo", true, file);

                                    //INTENSIDAD HORARIA
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtIntHora']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtIntHora']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtIntHora']", Intensidad);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Intensidad", true, file);

                                    //FECHA INICIAL
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", FechaInicial);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Fecha Inicial", true, file);

                                    //FECHA FINAL
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']", FechaFinal);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Fecha Final", true, file);

                                    //VALOR CURSO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtValCurs']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtValCurs']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValCurs']", Valor);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Valor", true, file);

                                    //ENTIDAD CAPACITORIA
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtEntCapa']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtEntCapa']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtEntCapa']", Entidad);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Entidad", true, file);

                                    //BUSCAR EMPLEADO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    Thread.Sleep(3000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']", Empleado);
                                    Thread.Sleep(3000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("EMPLEADO", true, file);

                                    //SELECCIONAR EMPLEADO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl03_Lnkbfiltro']");
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl03_Lnkbfiltro']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("EMPLEADO SELECCIONADO", true, file);

                                    Thread.Sleep(3000);
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    Thread.Sleep(3000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    Thread.Sleep(2000);
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    selenium.Screenshot("Archivo adjunto", true, file);

                                    //APLICAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(3000);

                                    //ACEPTAR ALERTA
                                    selenium.AcceptAlert();
                                    Thread.Sleep(3000);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("REGISTRO EXITOSO", true, file);
                                    Thread.Sleep(5000);
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
        public void FD_FormaciónDesarrolloEvaluaciónCursos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloEvaluaciónCursos")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba
                                rows["Respuesta1"].ToString().Length != 0 && rows["Respuesta1"].ToString() != null &&
                                rows["Respuesta2"].ToString().Length != 0 && rows["Respuesta2"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Respuesta1 = rows["Respuesta1"].ToString();
                                string Respuesta2 = rows["Respuesta2"].ToString();

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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }

                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {
                                        string actualizar = $"update FD_DPLIN set CON_FIRM='S',ASI_STIO='S', DPL_APTO='S' where IDE_NTIF='51880161' and RMT_PARA='73'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                        string BorrarEvaluacion = $"delete from fd_licev where cod_empl= 51880161";
                                        db.UpdateDeleteInsert(BorrarEvaluacion, database, user);
                                        string actualizar1 = $"UPDATE FD_DPLCU SET  IND_CALI = NULL WHERE COD_EMPR = '421' AND RMT_PARA = '73' AND RMT_PLCU = '1' AND  RMT_DEFI = '1'";
                                        db.UpdateDeleteInsert(actualizar1, database, user);



                                    }
                                    else
                                    {
                                        //PARAMETRIZACION PREVIA
                                        string actualizar = $"update FD_DPLIN set DPL_APTO='S',CON_FIRM='S',ASI_STIO='S' where IDE_NTIF='123' and RMT_PARA='178'";
                                        db.UpdateDeleteInsert(actualizar, database, user);  
                                        string BorrarEvaluacion = $"delete from fd_licev where cod_empl= 123";
                                        db.UpdateDeleteInsert(BorrarEvaluacion, database, user);
                                        string actualizar1 = $"UPDATE FD_DPLCU SET  IND_CALI = NULL WHERE COD_EMPR = '9' AND RMT_PARA = '178' AND RMT_PLCU = '1' AND RMT_DEFI = '8'";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

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

                                    //INGRESO A MIS CURSOS/INSCRIPCION
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Evaluación de Cursos')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A EVALUACION CURSOS", true, file);

                                    //DETALLE
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgFdPlcur_ctl03_LinkButton1']/i");
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdPlcur_ctl03_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle", true, file);

                                    //DETALLE CURSO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgFdDplcu_ctl03_LinkButton1']/i");
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdDplcu_ctl03_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle curso", true, file);

                                    //ACEPTAR EVALUACION
                                    selenium.Scroll(" //input[@id='ctl00_ContentPopapModel_Aceptar']");
                                    selenium.Click(" //input[@id='ctl00_ContentPopapModel_Aceptar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Confirmar", true, file);

                                    selenium.Click("//*[@id='ctl00_ContentPopapModel_dtgPreguntas_ctl02_SiguientePX']");
                                    ////RESPUESTA 1
                                    //selenium.Scroll("//table[@id='ctl00_ContentPopapModel_dtgPreguntas']/tbody/tr[2]/td[4]/select");
                                    //selenium.SelectElementByName("//table[@id='ctl00_ContentPopapModel_dtgPreguntas']/tbody/tr[2]/td[4]/select", Respuesta2);
                                    //Thread.Sleep(2000);
                                    //selenium.Screenshot("Respuesta 1", true, file);
                                    //selenium.Click("//*[@id='ctl00_ContentPopapModel_dtgPreguntas']/tbody/tr[2]/td[5]/a");
                                    //Thread.Sleep(2000);

                                    //selenium.Scroll("//select[@id='ctl00_ContentPopapModel_dtgPreguntas_ctl02_ddlRespuestas']");
                                    //selenium.SelectElementByName("//select[@id='ctl00_ContentPopapModel_dtgPreguntas_ctl02_ddlRespuestas']", Respuesta1);
                                    //Thread.Sleep(2000);
                                    //selenium.Screenshot("Respuesta 2", true, file);
                                    //selenium.Click("//*[@id='ctl00_ContentPopapModel_dtgPreguntas']/tbody/tr[2]/td[5]/a");
                                    //Thread.Sleep(2000);

                                    //selenium.Scroll("//*[@id='ctl00_ContentPopapModel_AceptarFin']");
                                    //selenium.Click("//*[@id='ctl00_ContentPopapModel_AceptarFin']");
                                    //Thread.Sleep(2000);
                                    selenium.ScrollTo("0", "900");
                                    selenium.Screenshot("Resultados", true, file);

                                    Thread.Sleep(5000);
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
        public void FD_FormaciónDesarrolloRevisiónEficaciaCapacitación()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloRevisiónEficaciaCapacitación")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null 


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }

                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {
                                        string actualizar = $"update FD_DPLIN set ASI_STIO='S' where IDE_NTIF='51880161' and RMT_PARA='73'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                        string actualizar1 = $"update FD_DPLIN set EVA_EFIC = null, RES_EFIC = null, OBS_EFIC = null, FEC_EFIC = null where IDE_NTIF = '51880161' and RMT_PARA = '73'";
                                        db.UpdateDeleteInsert(actualizar1, database, user);



                                    }
                                    else
                                    {
                                        //PARAMETRIZACION PREVIA
                                        string actualizar = $"update FD_DPLIN set ASI_STIO='S' where IDE_NTIF='123' and RMT_PARA='178'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                        string actualizar1 = $"update FD_DPLIN set EVA_EFIC = null, RES_EFIC = null, OBS_EFIC = null, FEC_EFIC = null where IDE_NTIF = '123' and RMT_PARA = '178'";
                                        db.UpdateDeleteInsert(actualizar1, database, user);
                                        


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

                                    //ROL LIDER
                                    selenium.Click("//*[@id='pLider']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    //FORMACION
                                    selenium.Click("//div/ul/li[3]/a");
                                    Thread.Sleep(2000);
                                    selenium.Click("(//a[contains(@href, 'frmLiFdRegefiL.aspx')])[2]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Medicion Eficacia", true, file);
                                    if (database == "SQL")
                                    {
                                        //REVISION EFICACIA DETALLE
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdPlcur_ctl04_LinkButton1']/i");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Detalle Eficacia", true, file);
                                    }
                                    else
                                    {
                                        //REVISION EFICACIA DETALLE
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdPlcur_ctl03_LinkButton1']/i");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Detalle Eficacia", true, file);
                                    }
                                    
                                    //EFICACIA DETALLE 2
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdDplse_ctl03_LinkButton1']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Medicion Eficacia", true, file);
                                    //RESPONDER MEDICION
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlregae']", "SI");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Mejoro Rendimiento", true, file);
                                    //OBSERVACIONES
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValM1_txtTexto']", "PRUEBAS CALIDAD");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    //APLICAR
                                    selenium.ScrollTo("0", "50");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnAplicar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Medicion Aplicada", true, file);

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(5000);
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
        public void FD_FormaciónDesarrolloReporteAsistenciaCurso()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloReporteAsistenciaCurso")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }

                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {
                                        string actualizar = $"update FD_DPLIN set ASI_STIO='S' where IDE_NTIF='51880161' and RMT_PARA='73'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                    }
                                    else
                                    {
                                        //PARAMETRIZACION PREVIA
                                        string actualizar = $"update FD_DPLIN set ASI_STIO='S' where IDE_NTIF='123' and RMT_PARA='178'";
                                        db.UpdateDeleteInsert(actualizar, database, user);


                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    string username = Environment.UserName;
                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    File.Delete("C:/Users/" + username + "/Downloads/AsistenciaCurso.pdf");

                                    //INGRESAR MIS CURSOS
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(500);
                                    if (database == "SQL")
                                    {
                                        selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[7]/a");
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[4]/ul/li[7]/a");
                                    }
                                    else
                                    {
                                        selenium.Scroll("//*[@id='MenuContex']/div[2]/div[1]/ul/li[4]/ul/li[11]/a");
                                        selenium.Click("//*[@id='MenuContex']/div[2]/div[1]/ul/li[4]/ul/li[11]/a");
                                    }

                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Reportes de Asistencia a Cursos", true, file);

                                    //DETALLE
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFdDplse_ctl03_lbSelDetalle']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle Seleccionado", true, file);

                                    //REPORTE
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnReporte']");
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnReporte']");
                                    Thread.Sleep(2000);
                                    Screenshot("Reporte Asistencia Generado", true, file);

                                    //VENTANA EMERGENTE
                                    String mainWin = selenium.MainWindow();
                                    String modalWin = selenium.PopupWindow();
                                    selenium.ChangeWindow(modalWin);
                                    Thread.Sleep(5000);
                                    Screenshot("Reporte Asistencia Generado", true, file);
                                    selenium.MaximizeWindow();
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Reporte Asistencia Cursos", true, file);
                                    Thread.Sleep(5000);
                                    //PDF IMPRIMIR
                                    selenium.Click("//*[@id='ctl00_btnImprimir']");
                                    Thread.Sleep(20000);
                                    Screenshot("IMPRIMIR PDF", true, file);
                                    //GUARDAR PDF
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }

                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{DOWN}");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);

                                    for (int i = 0; i < 6; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("AsistenciaCurso");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    selenium.ChangeWindow(mainWin);
                                    Thread.Sleep(5000);
                                    selenium.Close();

                                    //ABRIR PDF
                                    string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/" + username + "/Downloads/AsistenciaCurso.pdf");
                                    Process.Start(pdfPath);
                                    Thread.Sleep(6000);
                                    Screenshot("PDF ABIERTO", true, file);

                                    Thread.Sleep(2000);
                                   
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
                                    Thread.Sleep(5000);
                                    KillProcesos("Acrobat.exe");
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
        public void FD_FormaciónDesarrolloIngresoEvaluacionCursos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloIngresoEvaluacionCursos")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string calificacion = rows["Calificacion"].ToString();

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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }

                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {
                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='A' WHERE RMT_PARA='73'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
                                        string actualizar = $"update FD_DPLIN set CAL_SESI = null where IDE_NTIF = '51880161' and RMT_PARA = '73'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                    }
                                    else
                                    {
                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='A' WHERE RMT_PARA='178'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
                                        string actualizar = $"update FD_DPLIN set CAL_SESI = null where IDE_NTIF = '123' and RMT_PARA = '178'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                        
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
                                    //MIS CURSOS/INGRESO EVALUACION CURSOS
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'EF Ingreso de Evaluación de cursos')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'EF Ingreso de Evaluación de cursos')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Evaluación de cursos", true, file);
                                    if (selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_dtgFdClsesL_ctl03_LinkButton1']/i"))
                                    {
                                        //DETALLE
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdClsesL_ctl03_LinkButton1']/i");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Evaluación de cursos", true, file);
                                        //DETALLE 2
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdClsesD_ctl03_LinkButton1']/i");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Evaluación de cursos", true, file);
                                        //CALIFICACION
                                        selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_dtgFdClses_ctl02_txtVal_CaliPos']", calificacion);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Calificación Ingresada", true, file);
                                        //APLICAR
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_btnGuardar']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Evaluación aplicada", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Close();

                                    }
                                    else
                                    {
                                        Assert.Fail("No hay evaliaviones de cursos disponibles");
                                    }


                                    if (database == "ORA")
                                    {
                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='P' WHERE RMT_PARA='73'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
                                    }
                                    else
                                    {
                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='P' WHERE RMT_PARA='178'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
                                    }
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
        public void FD_FormaciónDesarrolloIngresoAsistenciaCursos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloIngresoAsistenciaCursos")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string asistencia = rows["Asistencia"].ToString();

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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }

                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {
                                        string actualizar = $"update FD_DPLIN set ASI_STIO = null where IDE_NTIF = '51880161' and RMT_PARA = '73'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='A' WHERE RMT_PARA='73'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
                                        

                                    }
                                    else
                                    {
                                        string actualizar = $"update FD_DPLIN set ASI_STIO = null where IDE_NTIF = '123' and RMT_PARA = '178'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='A' WHERE RMT_PARA='178'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
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
                                    //MIS CURSOS/INGRESO ASISTENCIA CURSOS
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Ingreso de Asistencia a Cursos')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Ingreso de Asistencia a Cursos')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Asistencia de cursos", true, file);

                                    if (selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_dtgFdvalasiL_ctl03_LinkButton1']/i"))
                                    {
                                        //DETALLE
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdvalasiL_ctl03_LinkButton1']/i");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Asistencia de cursos", true, file);
                                        //DETALLE 2
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdvalasid_ctl03_LinkButton1']/i");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Asistencia de cursos", true, file);
                                        //CALIFICACION
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_dtgFdvalasi_ctl02_ddlAsisSesion']", asistencia);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Asistencia Ingresada", true, file);
                                        //APLICAR
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_btnGuardar']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Asistencia confirmada", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Close();
                                    }
                                    else
                                    {
                                        Assert.Fail("No hay cursos para ingresar asistencia");
                                    }
                                   
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

                                    if (database == "ORA")
                                    {
                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='P' WHERE RMT_PARA='73'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
                                    }
                                    else
                                    {
                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='P' WHERE RMT_PARA='178'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);
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
        public void FD_FormaciónDesarrolloDocumentosCursos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloDocumentosCursos")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null 


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                    }

                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {
                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='A' WHERE RMT_PARA='73'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);


                                    }
                                    else
                                    {
                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='A' WHERE RMT_PARA='178'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);

                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string username = Environment.UserName;
                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               
                                    File.Delete(@"C:\Users\" + username + @"\Downloads\ArchivoArchivoPrueba.PDF");
                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A MIS CURSOS/INSCRIPCION
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(@href, 'frmFdPlcurLd.aspx')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Documentos Sesiones", true, file);

                                    //DETALLE
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgFdPlcur_ctl03_lbSelDetalle']/i");
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdPlcur_ctl03_lbSelDetalle']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle", true, file);

                                    //DOCUMENTOS SESIONES
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgFdDplse_ctl03_LinkButton1']/i");
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdDplse_ctl03_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle", true, file);

                                    //CONTENIDO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocu_ctl02_LinkButton3']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Documento a descargar", true, file);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocu_ctl02_LinkButton3']/i");
                                    Thread.Sleep(2000);
                                    Screenshot("Archivo descargado", true, file);

                                    //VENTANA EMERGENTE
                                    String mainWin = selenium.MainWindow();
                                    String modalWin = selenium.PopupWindow();
                                    Thread.Sleep(3000);
                                    selenium.ChangeWindow(modalWin);
                                    Thread.Sleep(3000);
                                    selenium.Close();
                                    selenium.ChangeWindow(mainWin);
                                    Thread.Sleep(5000);
                                    selenium.Close();
                                   
                                    string pdfPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\ArchivoArchivoPrueba.PDF");
                                    Process.Start(pdfPath);
                                    Thread.Sleep(6000);
                                    Screenshot("PDF ABIERTO", true, file);
                                    Thread.Sleep(5000);

                                    Thread.Sleep(5000);
                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {
                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='P' WHERE RMT_PARA='73'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);


                                    }
                                    else
                                    {
                                        string BorrarREgistro = $"update FD_PLCUR SET EST_CURS='P' WHERE RMT_PARA='178'";
                                        db.UpdateDeleteInsert(BorrarREgistro, database, user);

                                    }
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
                                    KillProcesos("Acrobat.exe");
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
        public void FD_FormaciónDesarrolloHistorialCapacitaciones()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloHistorialCapacitaciones")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba
                                rows["Fecha1"].ToString().Length != 0 && rows["Fecha1"].ToString() != null &&
                                rows["Fecha2"].ToString().Length != 0 && rows["Fecha2"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Fecha1 = rows["Fecha1"].ToString();
                                string Fecha2 = rows["Fecha2"].ToString();

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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
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
                                    string username = Environment.UserName;
                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    File.Delete("C:/Users/" + username + "/Downloads/HistoricoCursos.pdf");
                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A MIS CURSOS/INSCRIPCION
                                    selenium.Click("//a[contains(.,'MIS CURSOS')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Historial de Cursos-Capacitaciones')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("HISTORIAL CURSOS", true, file);

                                    //FECHA1
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_KCtrlFechaIni_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFechaIni_txtFecha']", Fecha1);
                                    Thread.Sleep(2000);

                                    //FECHA 2
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_KCtrlFechaFin_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFechaFin_txtFecha']", Fecha2);
                                    Thread.Sleep(2000);

                                    //ACEPTAR
                                    selenium.Screenshot("FECHAS", true, file);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnAceptar']");
                                    Thread.Sleep(2000);

                                    //HISTORICO
                                    selenium.Screenshot("REGISTRO HISTORIAL", true, file);
                                    Thread.Sleep(2000);

                                    //IMPRIMIR
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_btnimprimir']");
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_btnimprimir']");
                                    Thread.Sleep(2000);
                                    Screenshot("Histórico", true, file);
                                    //VENTANA 
                                    String mainWin = selenium.MainWindow();
                                    String modalWin = selenium.PopupWindow();
                                    Thread.Sleep(3000);
                                    selenium.ChangeWindow(modalWin);
                                    Thread.Sleep(5000);
                                    Screenshot("Reporte", true, file);
                                    selenium.MaximizeWindow();
                                    Thread.Sleep(5000);
                                    selenium.Click("//*[@id='Imprimir']");
                                    Thread.Sleep(5000);
                                    Screenshot("Reporte", true, file);
                                    //GUARDAR PDF
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }

                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{DOWN}");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);

                                    for (int i = 0; i < 4; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("HistoricoCursos");
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);

                                    //ABRIR PDF
                                    string pdfPath = Path.Combine(Application.StartupPath, "C:/Users/" + username + "/Downloads/HistoricoCursos.pdf");
                                    Process.Start(pdfPath);
                                    Thread.Sleep(6000);
                                    Screenshot("PDF ABIERTO", true, file);
                                    selenium.Close();
                                    selenium.ChangeWindow(mainWin);
                                    Thread.Sleep(5000);
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
                                    KillProcesos("Acrobat.exe");
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
        public void FD_FormaciónDesarrolloNecesidadesFormaciónAprobar()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloNecesidadesFormaciónAprobar")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba
                                rows["JefeUser"].ToString().Length != 0 && rows["JefeUser"].ToString() != null &&
                                rows["JefePass"].ToString().Length != 0 && rows["JefePass"].ToString() != null &&
                                 rows["Registro"].ToString().Length != 0 && rows["Registro"].ToString() != null &&
                                rows["Requerimiento"].ToString().Length != 0 && rows["Requerimiento"].ToString() != null &&
                                rows["Especificacion"].ToString().Length != 0 && rows["Especificacion"].ToString() != null &&
                                rows["Curso"].ToString().Length != 0 && rows["Curso"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();
                                string Registro = rows["Registro"].ToString();
                                string Requerimiento = rows["Requerimiento"].ToString();
                                string Especificacion = rows["Especificacion"].ToString();
                                string Curso = rows["Curso"].ToString();

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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
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

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {

                                        string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarSolicitud, database, user);
                                        string eliminarNecesidad = $"Delete from FD_NEFOR where COD_EMPL ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarNecesidad, database, user);
                                        string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{JefeUser}'";
                                        db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    }
                                    else
                                    {
                                       
                                        string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarSolicitud, database, user);
                                        string eliminarNecesidad = $"Delete from FD_NEFOR where COD_EMPL ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarNecesidad, database, user);
                                        string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{JefeUser}'";
                                        db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    }

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A MIS CURSOS/NECESIDADES FORMACION
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Mis necesidades de Formación')]");
                                    selenium.Click("//a[contains(.,'Mis necesidades de Formación')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A NECESIDADES DE FORMACION", true, file);

                                    //NUEVO
                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    Thread.Sleep(3000);

                                    //REGISTRO
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomRegi']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomRegi']", Registro);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro", true, file);
                                    //REQUERIMIENTO
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomRequ']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomRequ']", Requerimiento);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Requerimiento", true, file);
                                    //ESPECIFICACION
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomEspe']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomEspe']", Especificacion);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Especificacion", true, file);
                                    //CURSOS
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomCurs']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomCurs']", Curso);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Curso", true, file);
                                    //APLICAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(5000);

                                    //ACEPTAR ALERTA
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);

                                    //VERIFICAR REGISTRO
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Mis necesidades de Formación')]");
                                    selenium.Click("//a[contains(.,'Mis necesidades de Formación')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("REGISTRO EXITOSO", true, file);

                                    selenium.Close();

                                    //-----------------------------------------lider----------------------------------------------------------------
                                    //LOGIN
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A ROL LIDER
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                        Thread.Sleep(4000);
                                        selenium.Screenshot("ROL LIDER", true, file);
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'Lider')]");
                                        Thread.Sleep(4000);
                                        selenium.Screenshot("LIDER", true, file);
                                    }

                                    //INGRESO A FORMACION YU DESARROLLO
                                    selenium.Click("//a[contains(.,'FORMACIÓN Y DESARROLLO')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Aprobación de N. Formación')]");
                                    selenium.Click("//a[contains(.,'Aprobación de N. Formación')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A FORMACION Y DESARRLLO", true, file);

                                    //REGISTRO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1']/i");
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdNeforAJ_ctl03_LinkButton1']/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro", true, file);

                                    //APROBAR
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnAprobar']");
                                    selenium.Screenshot("Aprobar", true, file);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnAprobar']");
                                    Thread.Sleep(3000);

                                    //ACEPTAR ALERTA
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);

                                    //ACEPTAR ALERTA
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);

                                    selenium.Close();

                                    //---------------------Colaborador-------------------------------------------------

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A MIS CURSOS/NECESIDADES FORMACION
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Mis necesidades de Formación')]");
                                    selenium.Click("//a[contains(.,'Mis necesidades de Formación')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("NECESIDAD APORBADA", true, file);
                                    Thread.Sleep(5000);
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
        public void FD_FormaciónDesarrolloNecesidadesFormaciónAprobarGH()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_FD.FD_FormaciónDesarrolloNecesidadesFormaciónAprobarGH")
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
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba
                                rows["Registro"].ToString().Length != 0 && rows["Registro"].ToString() != null &&
                                rows["Requerimiento"].ToString().Length != 0 && rows["Requerimiento"].ToString() != null &&
                                rows["Especificacion"].ToString().Length != 0 && rows["Especificacion"].ToString() != null &&
                                rows["Curso"].ToString().Length != 0 && rows["Curso"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Registro = rows["Registro"].ToString();
                                string Requerimiento = rows["Requerimiento"].ToString();
                                string Especificacion = rows["Especificacion"].ToString();
                                string Curso = rows["Curso"].ToString();

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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewver/".ToLower())
                                    {
                                        database = "SQL";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
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

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //PARAMETRIZACION PREVIA
                                    if (database == "ORA")
                                    {

                                        string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarSolicitud, database, user);
                                        string eliminarNecesidad = $"Delete from FD_NEFOR where COD_EMPL ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarNecesidad, database, user);

                                    }
                                    else
                                    {

                                        string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='F' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarSolicitud, database, user);
                                        string eliminarNecesidad = $"Delete from FD_NEFOR where COD_EMPL ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarNecesidad, database, user);

                                    }

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A MIS CURSOS/NECESIDADES FORMACION
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Mis necesidades de Formación')]");
                                    selenium.Click("//a[contains(.,'Mis necesidades de Formación')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A NECESIDADES DE FORMACION", true, file);

                                    //NUEVO
                                    selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    Thread.Sleep(3000);

                                    //REGISTRO
                                    selenium.Click("//select[@id='ctl00_ContenidoPagina_ddlNomRegi']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomRegi']", Registro);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro", true, file);
                                    //REQUERIMIENTO
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomRequ']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomRequ']", Requerimiento);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Requerimiento", true, file);
                                    //ESPECIFICACION
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomEspe']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomEspe']", Especificacion);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Especificacion", true, file);
                                    //CURSOS
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomCurs']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomCurs']", Curso);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Curso", true, file);
                                    //APLICAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(5000);

                                    //ACEPTAR ALERTA
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);

                                    //VERIFICAR REGISTRO
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Mis necesidades de Formación')]");
                                    selenium.Click("//a[contains(.,'Mis necesidades de Formación')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("REGISTRO EXITOSO", true, file);



                                    //-----------------------------------------lider----------------------------------------------------------------


                                    //INGRESO A ROL RRHH
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(.,'Rol RRHH')]");
                                        Thread.Sleep(4000);
                                        selenium.Screenshot("ROL RRHH", true, file);
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'GESTION HUMANA')]");
                                        Thread.Sleep(4000);
                                        selenium.Screenshot("Gestion Humana", true, file);
                                    }

                                    //INGRESO A FORMACION YU DESARROLLO
                                    selenium.Click("//a[contains(.,'FORMACIÓN Y DESARROLLO')]");
                                    Thread.Sleep(2000);
                                    if (database == "SQL")
                                    {
                                        selenium.Scroll("//a[contains(.,'Solicitudes de Formación Pendientes por Aprobar')]");
                                        selenium.Click("//a[contains(.,'Solicitudes de Formación Pendientes por Aprobar')]");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("INGRESO A FORMACION Y DESARRLLO", true, file);

                                        //REGISTRO
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFdNefor_ctl03_LinkButton1']");
                                        Thread.Sleep(3000);
                                        selenium.Screenshot("Registro", true, file);

                                        //DETALLE
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdNefor_ctl02_LinkButton1']/i");
                                        Thread.Sleep(3000);
                                        selenium.Screenshot("Solicitud por aprobar", true, file);
                                    }
                                    else
                                    {
                                        selenium.Scroll("//li[5]/ul/li[4]/a");
                                        selenium.Click("//li[5]/ul/li[4]/a");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("INGRESO A FORMACION Y DESARRLLO", true, file);

                                        //REGISTRO
                                        selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgFdNefor_ctl10_LinkButton1']/i");
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdNefor_ctl10_LinkButton1']/i");
                                        Thread.Sleep(3000);
                                        selenium.Screenshot("Registro", true, file);

                                        //DETALLE
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgFdNefor_ctl02_LinkButton1']/i");
                                        Thread.Sleep(3000);
                                        selenium.Screenshot("Solicitud por aprobar", true, file);
                                    }
                                    //APROBAR
                                    
                                    selenium.Screenshot("Aprobar", true, file);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_Aprueba']");
                                    Thread.Sleep(3000);
                                    selenium.Click("//input[contains(@id,'btnEnviar')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Envia Correo Aprobación", true, file);
                                    //ACEPTAR ALERTA
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);

                                    //ACEPTAR ALERTA
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);

                                    Thread.Sleep(5000);
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
    }
}