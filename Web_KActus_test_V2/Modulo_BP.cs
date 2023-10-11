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
    public class Modulo_BP : FuncionesVitales
    {

        string Modulo = "Modulo_BP";
        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();

        public Modulo_BP()
        {

        }

        [TestMethod]
        public void BP_SolicitudServicioalClienteSelfService()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_SolicitudServicioalClienteSelfService")
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
                                //Datos Servicio al Cliente 
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                 rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["TipoSolicitud"].ToString().Length != 0 && rows["TipoSolicitud"].ToString() != null &&
                                rows["NumSolicitud"].ToString().Length != 0 && rows["NumSolicitud"].ToString() != null &&
                                rows["DescSolicitud"].ToString().Length != 0 && rows["DescSolicitud"].ToString() != null &&
                                rows["SecSoli"].ToString().Length != 0 && rows["SecSoli"].ToString() != null &&
                                rows["CodEmpresa"].ToString().Length != 0 && rows["CodEmpresa"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string CodPagina = rows["CodPagina"].ToString();
                                string url = rows["url"].ToString();
                                string TipoSolicitud = rows["TipoSolicitud"].ToString();
                                string NumSolicitud = rows["NumSolicitud"].ToString();
                                string DescSolicitud = rows["DescSolicitud"].ToString();
                                string SecSoli = rows["SecSoli"].ToString();
                                string CodEmpresa = rows["CodEmpresa"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    string Borrar1Tabla = $"DELETE FROM BI_SSOLQ WHERE ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(Borrar1Tabla, database, user);

                                    string Borrar1Tabla1 = $"DELETE FROM BI_DSOLQ WHERE ACT_USUA='Kactus'";
                                    db.UpdateDeleteInsert(Borrar1Tabla1, database, user);

                                    string Borrar2Tabla = $"DELETE FROM BI_SOLQU WHERE COD_REPO='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(Borrar2Tabla, database, user);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A SERVICIO AL CLIENTE INTERNO
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//a[contains(.,'SERVICIO CLIENTE INTERNO')]");
                                    selenium.Click("//a[contains(.,'SERVICIO CLIENTE INTERNO')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Servicio Cliente Interno')]");
                                    selenium.Screenshot("Servicio al cliente Interno", true, file);
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    selenium.Screenshot("Servicio al cliente Nuevo", true, file);


                                    Thread.Sleep(2000);

                                    //TIPO DE SOLICITUD
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodTiso')]", TipoSolicitud);
                                    Thread.Sleep(1000);
                                    //NUMERO DE SOLICITUD
                                    selenium.SendKeys("//input[contains(@id,'txtNUM_SOLI')]", NumSolicitud);
                                    selenium.Scroll("//textarea[contains(@id,'KCtrlTxtDesHech_txtTexto')]");
                                    Thread.Sleep(1000);
                                    //DESCRIPCION DE SOLICITUD
                                    selenium.SendKeys("//textarea[contains(@id,'KCtrlTxtDesHech_txtTexto')]", DescSolicitud);
                                    Thread.Sleep(1000);

                                    //GUARDAR
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_Agregar2']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Solicitud Guardada", true, file);
                                    Thread.Sleep(6000);
                                    selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Solicitud Enviada", true, file);

                                    //GENERACION CONSULTA VERIFICACION
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//a[contains(.,'SERVICIO CLIENTE INTERNO')]");
                                    selenium.Click("//a[contains(.,'SERVICIO CLIENTE INTERNO')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Servicio Cliente Interno')]");
                                    selenium.Screenshot("Servicio al cliente Interno", true, file);
                                    selenium.Screenshot("Se Genera consulta para verificar registro Guardado", true, file);

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
        public void BP_BeneficiosOrganizacionalesAprobaciónFamiliar()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_BeneficiosOrganizacionalesAprobaciónFamiliar")
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
                                //Datos Beneficios Organizacionales Aprobación
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["JefeUser"].ToString().Length != 0 && rows["JefeUser"].ToString() != null &&
                                rows["JefePass"].ToString().Length != 0 && rows["JefePass"].ToString() != null &&
                                rows["TipoAuxilio"].ToString().Length != 0 && rows["TipoAuxilio"].ToString() != null &&
                                rows["ValorSolicitado"].ToString().Length != 0 && rows["ValorSolicitado"].ToString() != null &&
                                rows["Observaciones"].ToString().Length != 0 && rows["Observaciones"].ToString() != null &&
                                rows["AsuntoCorreo"].ToString().Length != 0 && rows["AsuntoCorreo"].ToString() != null &&
                                rows["BodyCorreo"].ToString().Length != 0 && rows["BodyCorreo"].ToString() != null &&
                                rows["CodigoSolicitud"].ToString().Length != 0 && rows["CodigoSolicitud"].ToString() != null &&
                                rows["TipoApli"].ToString().Length != 0 && rows["TipoApli"].ToString() != null &&
                                rows["TipoDocumAdjunto"].ToString().Length != 0 && rows["TipoDocumAdjunto"].ToString() != null &&
                                rows["Ruta"].ToString().Length != 0 && rows["Ruta"].ToString() != null &&
                                rows["CodEmpresa"].ToString().Length != 0 && rows["CodEmpresa"].ToString() != null 


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
                                string TipoAuxilio = rows["TipoAuxilio"].ToString();
                                string ValorSolicitado = rows["ValorSolicitado"].ToString();
                                string Observaciones = rows["Observaciones"].ToString();
                                string AsuntoCorreo = rows["AsuntoCorreo"].ToString();
                                string BodyCorreo = rows["BodyCorreo"].ToString();
                                string CodigoSolicitud = rows["CodigoSolicitud"].ToString();
                                string TipoApli = rows["TipoApli"].ToString();
                                string TipoDocumAdjunto = rows["TipoDocumAdjunto"].ToString();
                                string CodEmpresa = rows["CodEmpresa"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string Modalidad = rows["Modalidad"].ToString();
                                string Programa = rows["Programa"].ToString();
                                string Calendario = rows["Calendario"].ToString();
                                string TipIntensidad = rows["TipIntensidad"].ToString();
                                string TipDocumento = rows["TipDocumento"].ToString();
                                string Periocidad = rows["Periocidad"].ToString();
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

                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    if (database == "SQL")
                                    {
                                        string EliminarConsecutivo = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo, database, user);

                                        string EliminarDetalle = $"Delete from BP_OTOBE where COD_EMPL ='{EmpleadoUser}' and COD_BENE ='4545'";
                                        db.UpdateDeleteInsert(EliminarDetalle, database, user);

                                        string EliminarConsecutivo1 = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{JefeUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo1, database, user);
                                    }
                                    else
                                    {
                                        string EliminarDetalle = $"Delete from BP_OTOBE where COD_BENE ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarDetalle, database, user);

                                        string EliminarConsecutivo = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo, database, user);

                                        string EliminarConsecutivo1 = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{JefeUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo1, database, user);

                                    }
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//span[contains(@id,'pColaborador')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    Thread.Sleep(1000);
                                    selenium.returnDriver().ExecuteScript("arguments[0].scrollIntoView(true);", selenium.returnDriver().FindElement(By.XPath("//div[@id='MenuContex']/div[2]/div/ul/li[3]/ul/li[15]/a")));
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/ul/li[15]/a");
                                    selenium.Screenshot("Beneficios Organizacionales", true, file);
                                    Thread.Sleep(200);

                                    //BENEFICIOS FAMILIAR
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_rbBenef_1')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Beneficio Familiar", true, file);
                                    if (database == "ORA")
                                    {
                                        //TIPO DE AUXILIO
                                        selenium.SelectElementByName("//select[contains(@id,'ddlNomTibe')]", TipoAuxilio);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Tipo auxilio", true, file);
                                        ////INGRESAR NUMERO CUOTAS                   
                                        //selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomCuot']", "2");
                                        //selenium.Screenshot("Entidad Beneficios", true, file);
                                        //selenium.Scroll("//select[@id='ctl00_ContenidoPagina_KCtrTipoDocumento2_ddlTIP_DOCU']");
                                        ////TIPO DE DOCUMENTO
                                        //selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_KCtrTipoDocumento2_ddlTIP_DOCU']", "5 COPIAS DE LA CÉDULA LEGIBLES AMPLIADAS AL 150%");
                                        //Thread.Sleep(2000);
                                        //selenium.Screenshot("Tipo Documento", true, file);
                                        //SELECCIONAR FAMILIAR
                                        selenium.Click("//input[@id='ctl00_ContenidoPagina_dtgBiFamil_ctl02_chcod']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Familiar", true, file);
                                        //INGRESAR VALOR SOLICITADO
                                        Thread.Sleep(2000);
                                        selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_dtgBiFamil_ctl02_ValSoli']", ValorSolicitado);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Valor Solicitado", true, file);
                                        //AGREGAR OBSERVACIONES
                                        selenium.Scroll("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]");
                                        selenium.SendKeys("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]", Observaciones);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Observaciones", true, file);
                                      
                                    }
                                    else
                                    {
                                        //TIPO DE AUXILIO
                                        selenium.SelectElementByName("//select[contains(@id,'ddlNomTibe')]", TipoAuxilio);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Tipo auxilio", true, file);
                                        //ENTIDAD
                                        selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_ddlEntBen']", "SENA");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Entidad", true, file);
                                        //MODALIDAD
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlClaSifi')]", Modalidad);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Modalidad", true, file);
                                        //PROGRAMA
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlProGram')]", Programa);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Programa", true, file);
                                        //CALENDARIO
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlCalEnda')]", Calendario);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Calendario", true, file);
                                        //TIPO DE INTENSIDAD
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlSemEstr')]", TipIntensidad);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Tipo intensidad", true, file);
                                        selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgBiFamil_ctl02_chcod']");
                                        Thread.Sleep(2000);
                                        //PERIOCIDAD
                                        selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_ddlPerIodi']",Periocidad);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Periocidad", true, file);
                                        ////SELECCIONAR FAMILIAR
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgBiFamil_ctl02_chcod']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Familiar", true, file);
                                        ////VALOR UNITARIO
                                        selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_dtgBiFamil_ctl02_ValSoli']", ValorSolicitado);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Valor solicitado", true, file);
                                        selenium.Scroll("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]");
                                        Thread.Sleep(2000);
                                        //OBSERVACIONES
                                        selenium.SendKeys("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]", Observaciones);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Observaciones", true, file);
                                    }
                                    
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("solicitud Guardada", true, file);
                                    Thread.Sleep(500);
                                    selenium.ChangeAuxWindow();
                                    if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                    {
                                        selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    }
                                    selenium.Screenshot("Solicitud Radicada", true, file);
                                    Thread.Sleep(500);
                                    selenium.ChangeMainWindow();
                                    Thread.Sleep(2000);
                                    //CONSULTAR TODO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Solicitud Registrada", true, file);
                                    selenium.Close();

                                    //APROBACION POR JEFE
                                    //Aprobación por parte del Jefe 
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    selenium.Click("//span[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[12]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[12]/a");
                                    Thread.Sleep(2000);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[12]/ul/li/a");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Aprobación Beneficios organizacionales", true, file);
                                    Thread.Sleep(500);

                                    if (selenium.ExistControl("//*[@id='tblBpBeoto']/tbody/tr/td[8]/a"))
                                    {
                                        selenium.Click("//*[@id='tblBpBeoto']/tbody/tr/td[8]/a");
                                        selenium.Screenshot("Seleccionar Registro", true, file);
                                        Thread.Sleep(2000);
                                        
                                        selenium.Screenshot("Seleccionar Registro para Aprobación", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_Aprueba')]");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Aprobación", true, file);
                                        Thread.Sleep(1000);
                                        selenium.SendKeys("//input[contains(@id,'txtAsuMail')]", AsuntoCorreo);
                                        selenium.SendKeys("//textarea[contains(@id,'txtCueMail')]", BodyCorreo);
                                        selenium.Screenshot("Información Correo Aprobación", true, file);
                                        Thread.Sleep(1000);
                                        selenium.Click("//div[@id='ctl00_pBotones']/div");
                                        Thread.Sleep(1000);
                                        selenium.Click("//input[contains(@id,'btnEnviar')]");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Envia Correo Aprobación", true, file);

                                    }
                                    else
                                    {
                                        Assert.Fail("ERROR: NO HAY BENEFICIOS ORGANIZACIONALES PENDIENTES POR APROBAR DE MIS COLABORADORES");
                                    }

                                    Thread.Sleep(3000);
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file,database);
                                    Thread.Sleep(6000);
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
        public void BP_BeneficiosOrganizacionalesAprobaciónColaborador()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_BeneficiosOrganizacionalesAprobaciónColaborador")
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
                                //Datos Beneficios Organizacionales Aprobación
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["JefeUser"].ToString().Length != 0 && rows["JefeUser"].ToString() != null &&
                                rows["JefePass"].ToString().Length != 0 && rows["JefePass"].ToString() != null &&
                                rows["TipoAuxilio"].ToString().Length != 0 && rows["TipoAuxilio"].ToString() != null &&
                                rows["ValorSolicitado"].ToString().Length != 0 && rows["ValorSolicitado"].ToString() != null &&
                                rows["Observaciones"].ToString().Length != 0 && rows["Observaciones"].ToString() != null &&
                                rows["AsuntoCorreo"].ToString().Length != 0 && rows["AsuntoCorreo"].ToString() != null &&
                                rows["BodyCorreo"].ToString().Length != 0 && rows["BodyCorreo"].ToString() != null &&
                                rows["CodigoSolicitud"].ToString().Length != 0 && rows["CodigoSolicitud"].ToString() != null &&
                                rows["TipoApli"].ToString().Length != 0 && rows["TipoApli"].ToString() != null &&
                                rows["TipoDocumAdjunto"].ToString().Length != 0 && rows["TipoDocumAdjunto"].ToString() != null &&
                                rows["Ruta"].ToString().Length != 0 && rows["Ruta"].ToString() != null &&
                                rows["CodEmpresa"].ToString().Length != 0 && rows["CodEmpresa"].ToString() != null


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
                                string TipoAuxilio = rows["TipoAuxilio"].ToString();
                                string ValorSolicitado = rows["ValorSolicitado"].ToString();
                                string Observaciones = rows["Observaciones"].ToString();
                                string AsuntoCorreo = rows["AsuntoCorreo"].ToString();
                                string BodyCorreo = rows["BodyCorreo"].ToString();
                                string CodigoSolicitud = rows["CodigoSolicitud"].ToString();
                                string TipoApli = rows["TipoApli"].ToString();
                                string TipoDocumAdjunto = rows["TipoDocumAdjunto"].ToString();
                                string CodEmpresa = rows["CodEmpresa"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string Modalidad = rows["Modalidad"].ToString();
                                string Programa = rows["Programa"].ToString();
                                string Calendario = rows["Calendario"].ToString();
                                string TipIntensidad = rows["TipIntensidad"].ToString();
                                string TipDocumento = rows["TipDocumento"].ToString();
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

                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    if (database == "ORA")
                                    {
                                        string EliminarDetalle = $"Delete from BP_OTOBE where COD_BENE ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarDetalle, database, user);

                                        string EliminarConsecutivo = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo, database, user);

                                        string EliminarConsecutivo1 = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{JefeUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo1, database, user);

                                    }else if (database == "SQL")
                                    {
                                        string EliminarDetalle = $"Delete from BP_OTOBE where COD_BENE ='{EmpleadoUser}' AND COD_TIBE='2'";
                                        db.UpdateDeleteInsert(EliminarDetalle, database, user);

                                        string EliminarConsecutivo = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo, database, user);

                                        string EliminarConsecutivo1 = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{JefeUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo1, database, user);

                                    }

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//span[contains(@id,'pColaborador')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    selenium.Screenshot("Mis Solicitudes", true, file);


                                    Thread.Sleep(1000);
                                    selenium.returnDriver().ExecuteScript("arguments[0].scrollIntoView(true);", selenium.returnDriver().FindElement(By.XPath("//div[@id='MenuContex']/div[2]/div/ul/li[3]/ul/li[15]/a")));
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/ul/li[15]/a");
                                    selenium.Screenshot("Beneficios Organizacionales", true, file);

                                    //BENEFICIOS

                                    selenium.Click("//input[contains(@id,'rbBenef_0')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Beneficio colaborador", true, file);
                                    //TIPO DE AUXILIO
                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomTibe')]", TipoAuxilio);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo auxilio", true, file);
                                    //INGRESAR ENTIDAD BENEFICIO  
                                    if (database == "ORA")
                                        {
                                            selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlEntBen')]", "Prueba");
                                            selenium.Screenshot("Entidad Beneficios", true, file);
                                         }   
                                     //INGRESAR VALOR SOLICITADO
                                     Thread.Sleep(2000);
                                     selenium.SendKeys("//input[contains(@id,'txtValSoli')]", ValorSolicitado);
                                     Thread.Sleep(2000);
                                     selenium.Screenshot("Valor Solicitado", true, file);
                                     //AGREGAR OBSERVACIONES
                                     selenium.Scroll("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]", Observaciones);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    //ADJUNTO
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(5000);
                                   
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Solicitud Guardada", true, file);
                                    Thread.Sleep(500);
                                    selenium.ChangeAuxWindow();
                                    if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                    {
                                        selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                        Thread.Sleep(2000);
                                    }
                                    selenium.Screenshot("Solicitud Radicada", true, file);
                                    Thread.Sleep(500);
                                    selenium.ChangeMainWindow();
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.Close();


                                    //APROBACION POR JEFE
                                    //Aprobación por parte del Jefe 
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    selenium.Click("//span[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[12]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[12]/a");
                                    Thread.Sleep(500);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[12]/ul/li/a");
                                    selenium.Screenshot("Aprobación Beneficios organizacionales", true, file);
                                    Thread.Sleep(500);

                                    if (selenium.ExistControl("//*[@id='tblBpBeoto']/tbody/tr/td[8]/a"))
                                    {
                                        selenium.Click("//*[@id='tblBpBeoto']/tbody/tr/td[8]/a");
                                        selenium.Screenshot("Seleccionar Registro", true, file);
                                        Thread.Sleep(2000);
                                        
                                        selenium.Screenshot("Seleccionar Registro para Aprobación", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_Aprueba')]");
                                        selenium.Screenshot("Aprobación", true, file);
                                        Thread.Sleep(1000);
                                        selenium.SendKeys("//input[contains(@id,'txtAsuMail')]", AsuntoCorreo);
                                        selenium.SendKeys("//textarea[contains(@id,'txtCueMail')]", BodyCorreo);
                                        selenium.Screenshot("Información Correo Aprobación", true, file);
                                        Thread.Sleep(1000);
                                        selenium.Click("//div[@id='ctl00_pBotones']/div");
                                        Thread.Sleep(1000);
                                        selenium.Click("//input[contains(@id,'btnEnviar')]");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Envia Correo Aprobación", true, file);

                                    }
                                    else
                                    {
                                        Assert.Fail("ERROR: NO HAY BENEFICIOS ORGANIZACIONALES PENDIENTES POR APROBAR DE MIS COLABORADORES");
                                    }
                                    Thread.Sleep(3000);
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(6000);
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
        public void BP_BeneficiosOrganizacionalesRechazoFamiliar()
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
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_BeneficiosOrganizacionalesRechazoFamiliar")
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
                                //Datos Beneficios Organizacionales Rechazo
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["JefeUser"].ToString().Length != 0 && rows["JefeUser"].ToString() != null &&
                                rows["JefePass"].ToString().Length != 0 && rows["JefePass"].ToString() != null &&
                                rows["TipoAuxilio"].ToString().Length != 0 && rows["TipoAuxilio"].ToString() != null &&
                                rows["ValorSolicitado"].ToString().Length != 0 && rows["ValorSolicitado"].ToString() != null &&
                                rows["TipoApli"].ToString().Length != 0 && rows["TipoApli"].ToString() != null &&
                                rows["TipoDocumAdjunto"].ToString().Length != 0 && rows["TipoDocumAdjunto"].ToString() != null &&
                                rows["Ruta"].ToString().Length != 0 && rows["Ruta"].ToString() != null &&
                                rows["CodEmpresa"].ToString().Length != 0 && rows["CodEmpresa"].ToString() != null
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
                                string TipoAuxilio = rows["TipoAuxilio"].ToString();
                                string ValorSolicitado = rows["ValorSolicitado"].ToString();
                                string TipoApli = rows["TipoApli"].ToString();
                                string TipoDocumAdjunto = rows["TipoDocumAdjunto"].ToString();
                                string CodEmpresa = rows["CodEmpresa"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string Modalidad = rows["Modalidad"].ToString();
                                string Programa = rows["Programa"].ToString();
                                string Calendario = rows["Calendario"].ToString();
                                string TipIntensidad = rows["TipIntensidad"].ToString();
                                string TipDocumento = rows["TipDocumento"].ToString();
                                string Observaciones = rows["Observaciones"].ToString();
                                string Periocidad= rows["Periocidad"].ToString();
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

                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //ELIMINAR REGISTROS PREVIOS
                                    if (database == "SQL")
                                    {
                                        string EliminarConsecutivo = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo, database, user);

                                        string EliminarDetalle = $"Delete from BP_OTOBE where COD_EMPL ='{EmpleadoUser}' and COD_BENE ='4545'";
                                        db.UpdateDeleteInsert(EliminarDetalle, database, user);

                                        string EliminarConsecutivo1 = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{JefeUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo1, database, user);
                                    }
                                    else
                                    {
                                        string EliminarDetalle = $"Delete from BP_OTOBE where COD_BENE ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarDetalle, database, user);

                                        string EliminarConsecutivo = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo, database, user);

                                        string EliminarConsecutivo1 = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{JefeUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo1, database, user);

                                    }


                                    //INICIO PRUEBA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    //BENEFICIOS ORGANIZACIONALES
                                    selenium.Click("//span[contains(@id,'pColaborador')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    Thread.Sleep(1000);
                                    selenium.returnDriver().ExecuteScript("arguments[0].scrollIntoView(true);", selenium.returnDriver().FindElement(By.XPath("//div[@id='MenuContex']/div[2]/div/ul/li[3]/ul/li[15]/a")));
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/ul/li[15]/a");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Beneficios Organizacionales", true, file);
                                    Thread.Sleep(200);
                                    //BENEFICIO FAMILIAR
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_rbBenef_1')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Beneficio Familiar", true, file);
                                    if (database == "ORA")
                                    {
                                        //TIPO DE AUXILIO
                                        selenium.SelectElementByName("//select[contains(@id,'ddlNomTibe')]", TipoAuxilio);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Tipo auxilio", true, file);
                                        ////INGRESAR NUMERO CUOTAS                   
                                        //selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomCuot']", "2");
                                        //selenium.Screenshot("Entidad Beneficios", true, file);
                                        //selenium.Scroll("//select[@id='ctl00_ContenidoPagina_KCtrTipoDocumento2_ddlTIP_DOCU']");
                                        ////TIPO DE DOCUMENTO
                                        //selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_KCtrTipoDocumento2_ddlTIP_DOCU']", "5 COPIAS DE LA CÉDULA LEGIBLES AMPLIADAS AL 150%");
                                        //selenium.Screenshot("Entidad Beneficios", true, file);
                                        //SELECCIONAR FAMILIAR
                                        selenium.Scroll("//input[@id='ctl00_ContenidoPagina_dtgBiFamil_ctl02_chcod']");
                                        Thread.Sleep(2000);
                                        selenium.Click("//input[@id='ctl00_ContenidoPagina_dtgBiFamil_ctl02_chcod']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Familiar", true, file);
                                        //INGRESAR VALOR SOLICITADO
                                        selenium.Scroll("//input[@id='ctl00_ContenidoPagina_dtgBiFamil_ctl02_ValSoli']");
                                        Thread.Sleep(2000);
                                        selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_dtgBiFamil_ctl02_ValSoli']", ValorSolicitado);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Valor Solicitado", true, file);
                                        //AGREGAR OBSERVACIONES
                                        selenium.Scroll("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]");
                                        Thread.Sleep(2000);
                                        selenium.SendKeys("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]", Observaciones);
                                        selenium.Screenshot("Observaciones", true, file);

                                    }
                                    else
                                    {
                                        //TIPO DE AUXILIO
                                        selenium.SelectElementByName("//select[contains(@id,'ddlNomTibe')]", TipoAuxilio);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Tipo auxilio", true, file);
                                        //ENTIDAD
                                        selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_ddlEntBen']", "SENA");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Entidad", true, file);
                                        //MODALIDAD
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlClaSifi')]", Modalidad);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Modalidad", true, file);
                                        //PROGRAMA
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlProGram')]", Programa);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Programa", true, file);
                                        //CALENDARIO
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlCalEnda')]", Calendario);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Calendario", true, file);
                                        //TIPO DE INTENSIDAD
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlSemEstr')]", TipIntensidad);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Tipo intensidad", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Scroll("//*[@id='ctl00_ContenidoPagina_ddlPerIodi']");
                                        Thread.Sleep(2000);
                                        //PERIOCIDAD
                                        selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_ddlPerIodi']", Periocidad);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Periocidad", true, file);
                                        ////SELECCIONAR FAMILIAR
                                        selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgBiFamil_ctl02_chcod']");
                                        Thread.Sleep(2000);
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgBiFamil_ctl02_chcod']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Familiar", true, file);
                                        ////VALOR UNITARIO
                                        selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_dtgBiFamil_ctl02_ValSoli']", ValorSolicitado);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Valor solicitado", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Scroll("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]");
                                        Thread.Sleep(2000);
                                        //OBSERVACIONES
                                        selenium.SendKeys("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]", Observaciones);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Observaciones", true, file);
                                    }
                                    // GUARDAR REGISTRO

                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Solicitud Guardada", true, file);
                                    Thread.Sleep(500);
                                    selenium.ChangeAuxWindow();
                                    if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                    {
                                        selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    }
                                    selenium.Screenshot("Solicitud Radicada", true, file);
                                    Thread.Sleep(500);
                                    selenium.ChangeMainWindow();
                                    Thread.Sleep(2000);
                                    //CONSULTAR TODOS
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Solicitud Registrada", true, file);
                                    selenium.Close();
                                    //APROBACION POR EL JEFE
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    selenium.Click("//span[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[12]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[12]/a");
                                    Thread.Sleep(500);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[12]/ul/li/a");
                                    selenium.Screenshot("Aprobación Beneficios organizacionales", true, file);
                                    Thread.Sleep(500);
                                    if (selenium.ExistControl("//*[@id='tblBpBeoto']/tbody/tr/td[8]/a"))
                                    {
                                        selenium.Click("//*[@id='tblBpBeoto']/tbody/tr/td[8]/a");
                                        selenium.Screenshot("Seleccionar Registro", true, file);
                                        Thread.Sleep(2000);
                                        
                                        selenium.Screenshot("Seleccionar Registro para Aprobación", true, file);
                                        Thread.Sleep(2000);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Rechazo Beneficio", true, file);
                                        selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_Rechaza')]");
                                        Thread.Sleep(2000);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Envia Correo Rechazo", true, file);
                                        selenium.Click("//input[contains(@id,'btnEnviar')]");
                                    }
                                    else
                                    {
                                        Assert.Fail("ERROR: NO HAY BENEFICIOS ORGANIZACIONALES PENDIENTES POR APROBAR DE MIS COLABORADORES");
                                    }

                                    Thread.Sleep(3000);
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(1000);
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
        public void BP_BeneficiosOrganizacionalesRechazoColaborador()
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
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_BeneficiosOrganizacionalesRechazoColaborador")
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
                                //Datos Beneficios Organizacionales Rechazo
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["JefeUser"].ToString().Length != 0 && rows["JefeUser"].ToString() != null &&
                                rows["JefePass"].ToString().Length != 0 && rows["JefePass"].ToString() != null &&
                                rows["TipoAuxilio"].ToString().Length != 0 && rows["TipoAuxilio"].ToString() != null &&
                                rows["ValorSolicitado"].ToString().Length != 0 && rows["ValorSolicitado"].ToString() != null &&
                                rows["TipoApli"].ToString().Length != 0 && rows["TipoApli"].ToString() != null &&
                                rows["TipoDocumAdjunto"].ToString().Length != 0 && rows["TipoDocumAdjunto"].ToString() != null &&
                                rows["Ruta"].ToString().Length != 0 && rows["Ruta"].ToString() != null &&
                                rows["CodEmpresa"].ToString().Length != 0 && rows["CodEmpresa"].ToString() != null
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
                                string TipoAuxilio = rows["TipoAuxilio"].ToString();
                                string ValorSolicitado = rows["ValorSolicitado"].ToString();
                                string TipoApli = rows["TipoApli"].ToString();
                                string TipoDocumAdjunto = rows["TipoDocumAdjunto"].ToString();
                                string CodEmpresa = rows["CodEmpresa"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string Modalidad = rows["Modalidad"].ToString();
                                string Programa = rows["Programa"].ToString();
                                string Calendario = rows["Calendario"].ToString();
                                string TipIntensidad = rows["TipIntensidad"].ToString();
                                string TipDocumento = rows["TipDocumento"].ToString();
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

                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //ELIMINAR REGISTROS PREVIOS

                                    if (database == "ORA")
                                    {
                                        string EliminarDetalle = $"Delete from BP_OTOBE where COD_BENE ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarDetalle, database, user);

                                        string EliminarConsecutivo = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo, database, user);

                                        string EliminarConsecutivo1 = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{JefeUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo1, database, user);

                                    }
                                    else if (database == "SQL")
                                    {
                                        string EliminarDetalle = $"Delete from BP_OTOBE where COD_BENE ='{EmpleadoUser}' AND COD_TIBE='2'";
                                        db.UpdateDeleteInsert(EliminarDetalle, database, user);

                                        string EliminarConsecutivo = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo, database, user);

                                        string EliminarConsecutivo1 = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{JefeUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo1, database, user);

                                    }


                                    //INICIO PRUEBA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    //MIS SOLICITUDES
                                    selenium.Click("//span[contains(@id,'pColaborador')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Mis Solicitudes", true, file);
                                    Thread.Sleep(2000);
                                    selenium.returnDriver().ExecuteScript("arguments[0].scrollIntoView(true);", selenium.returnDriver().FindElement(By.XPath("//div[@id='MenuContex']/div[2]/div/ul/li[3]/ul/li[15]/a")));
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/ul/li[15]/a");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Beneficios Organizacionales", true, file);
                                    Thread.Sleep(200);

                                    //BENEFICIO
                                    selenium.Click("//input[contains(@id,'rbBenef_0')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Beneficio Colaborador", true, file);
                                    //INGRESO TIPO DE AUXILIO
                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomTibe')]", TipoAuxilio);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo auxilio", true, file);
                                    //INGRESAR ENTIDAD BENEFICIO
                                    if (database == "ORA")
                                    {
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlEntBen')]", "Prueba");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Entidad", true, file);
                                    }
                                    //INGRESAR VALOR SOLICITADO
                                    selenium.Scroll("//input[contains(@id,'txtValSoli')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtValSoli')]", ValorSolicitado);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Valor", true, file);
                                    selenium.Scroll("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]");
                                    //AGREGAR OBSERVACIONES
                                    selenium.Scroll("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]", "PRUEBA");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    //ADJUNTO
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(5000);
                    
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    // GUARDAR REGISTRO
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Solicitud Guardada", true, file);
                                    Thread.Sleep(500);
                                    selenium.ChangeAuxWindow();
                                    if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                    {
                                        selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    }
                                    selenium.Screenshot("Solicitud Radicada", true, file);
                                    Thread.Sleep(500);
                                    selenium.ChangeMainWindow();
                                    Thread.Sleep(2000);
                                    //CONSULTAR TODOS
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Solicitud Registrada", true, file);
                                    selenium.Close();

                                    //APROBACION POR EL JEFE
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    selenium.Click("//span[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[12]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[12]/a");
                                    Thread.Sleep(500);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[12]/ul/li/a");
                                    selenium.Screenshot("Aprobación Beneficios organizacionales", true, file);
                                    Thread.Sleep(500);

                                    if (selenium.ExistControl("//*[@id='tblBpBeoto']/tbody/tr/td[8]/a"))
                                    {
                                        selenium.Click("//*[@id='tblBpBeoto']/tbody/tr/td[8]/a");
                                        selenium.Screenshot("Seleccionar Registro", true, file);
                                        Thread.Sleep(2000);
                                       
                                        selenium.Screenshot("Seleccionar Registro para Aprobación", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_Rechaza')]");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Rechazo Beneficio", true, file);
                                        Thread.Sleep(1000);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Envia Correo Rechazo", true, file);
                                        selenium.Click("//input[contains(@id,'btnEnviar')]");
                                        Thread.Sleep(2000);
                                    }
                                    else
                                    {
                                        Assert.Fail("ERROR: NO HAY BENEFICIOS ORGANIZACIONALES PENDIENTES POR APROBAR DE MIS COLABORADORES");
                                    }

                                    Thread.Sleep(3000);
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(1000);
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
        public void BP_SolicitudPréstamosConIntereses()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_SolicitudPréstamosConIntereses")
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

                            if (rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["EmailRecepcion"].ToString().Length != 0 && rows["EmailRecepcion"].ToString() != null &&
                                rows["PassEmailRecepcion"].ToString().Length != 0 && rows["PassEmailRecepcion"].ToString() != null &&
                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["Observa"].ToString().Length != 0 && rows["Observa"].ToString() != null &&
                                rows["DesPrestamo"].ToString().Length != 0 && rows["DesPrestamo"].ToString() != null &&
                                rows["PagoMens"].ToString().Length != 0 && rows["PagoMens"].ToString() != null &&
                                rows["Contenido"].ToString().Length != 0 && rows["Contenido"].ToString() != null &&
                                rows["CuotaMens"].ToString().Length != 0 && rows["CuotaMens"].ToString() != null &&
                                rows["NumCuotas"].ToString().Length != 0 && rows["NumCuotas"].ToString() != null &&
                                rows["Remodela"].ToString().Length != 0 && rows["Remodela"].ToString() != null &&
                                rows["ValInmueble"].ToString().Length != 0 && rows["ValInmueble"].ToString() != null &&
                                rows["Vigentes"].ToString().Length != 0 && rows["Vigentes"].ToString() != null &&
                                rows["Externas"].ToString().Length != 0 && rows["Externas"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["RmtSoli"].ToString().Length != 0 && rows["RmtSoli"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["ValCuota"].ToString().Length != 0 && rows["ValCuota"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string CodPagina = rows["CodPagina"].ToString();
                                string EmailRecepcion = rows["EmailRecepcion"].ToString();
                                string PassEmailRecepcion = rows["PassEmailRecepcion"].ToString();
                                string Descripcion = rows["Descripcion"].ToString();
                                string Observa = rows["Observa"].ToString();
                                string DesPrestamo = rows["DesPrestamo"].ToString();
                                string PagoMens = rows["PagoMens"].ToString();
                                string CuotaMens = rows["CuotaMens"].ToString();
                                string NumCuotas = rows["NumCuotas"].ToString();
                                string Remodela = rows["Remodela"].ToString();
                                string ValInmueble = rows["ValInmueble"].ToString();
                                string Vigentes = rows["Vigentes"].ToString();
                                string Externas = rows["Externas"].ToString();

                                string user = rows["user"].ToString();
                                string Contenido = rows["Contenido"].ToString();
                                string RmtSoli = rows["RmtSoli"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string ValCuota = rows["ValCuota"].ToString();
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

                                    List<string> errorsTest = new List<string>();
                                    List<string> errors = new List<string>();
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    if (database == "SQL")
                                    {
                                        //ELIMINAR REGISTRO
                                        string eliminarRegistro2 = $"DELETE FROM bp_dsoli where act_usua='202020' and cod_empr='9'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro = $"DELETE FROM bp_sopre where cod_empl='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);

                                    }
                                    else
                                    {
                                        string eliminarRegistro2 = $"DELETE FROM bp_dsoli where act_usua='193454' and cod_empr='421'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro = $"DELETE FROM bp_sopre where cod_empl='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                    }

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//a[contains(.,'Mis Solicitudes de Préstamos Sin Interes')]");
                                    selenium.Click("//a[contains(.,'Mis Solicitudes de Préstamos Sin Interes')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Préstamos Sin Interes", true, file);
                                    Thread.Sleep(1000);
                                    selenium.Click("//a[contains(@id,'ctl00_btnNuevo')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cerrar notificación", true, file);
                                    Thread.Sleep(3000);
                                    if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                    {
                                        selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    }
                                    selenium.Screenshot("MIS SOLICITUDES DE PRÉSTAMOS SIN INTERÉS", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_txtDedExte')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_txtDedExte')]");
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtDedExte')]", Externas);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Deudas Externas", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_txtComVige')]");
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_txtComVige')]");
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtComVige')]", Vigentes);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Compromisos Vigentes", true, file);
                                    selenium.Scroll("//select[contains(@id,'ctl00_ContenidoPagina_ddlNomPres')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlNomPres')]", DesPrestamo);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Descripcion Préstamo", true, file);
                                    Thread.Sleep(1500);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_txtValInmu')]");
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtValInmu')]", ValInmueble);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Valor inmueble", true, file);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_txtValRemo')]");
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtValRemo')]", Remodela);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Valor remodelación", true, file);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_txtValPres')]");
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtValPres')]", CuotaMens);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(3000);
                                
                                    selenium.Screenshot("Cuota mensual", true, file);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_txtNumCuot')]");
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtNumCuot')]", NumCuotas);
                                    Thread.Sleep(5000);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtNumCuot')]", NumCuotas);
                                    Thread.Sleep(5000);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Número cuotas", true, file);
                                    selenium.Scroll("//*[@id=\"ctl00_ContenidoPagina_txtValCuot\"]");
                                    selenium.SendKeys("//*[@id=\"ctl00_ContenidoPagina_txtValCuot\"]", ValCuota);
                                    Thread.Sleep(5000);
                                    selenium.SendKeys("//*[@id=\"ctl00_ContenidoPagina_txtValCuot\"]", ValCuota);
                                    Thread.Sleep(3000);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Valor cuota", true, file);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_txtPorCume')]");
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtPorCume')]", PagoMens);
                                    Thread.Sleep(5000);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtPorCume')]", PagoMens);
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Pago mensual", true, file);
                                    Thread.Sleep(5000);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtObsErva_txtTexto']");
                                    selenium.Click("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtObsErva_txtTexto']");
                                    Thread.Sleep(5000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtObsErva_txtTexto']", Observa);
                                    Thread.Sleep(5000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtObsErva_txtTexto']", Observa);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(10000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(10000);
                                    selenium.Screenshot("Registro", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//a[contains(.,'Mis Solicitudes de Préstamos Sin Interes')]");
                                    selenium.Click("//a[contains(.,'Mis Solicitudes de Préstamos Sin Interes')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Registro", true, file);

                                    Thread.Sleep(4000);
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database);
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

                                    Thread.Sleep(2000);
                             
                                    bandera = true;
                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();
                                        int StCount = Int32.Parse(SthCount);

                                        if (StCount > 0)
                                        {
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
        public void BP_SolicitudPréstamo()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_SolicitudPréstamo")
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

                            if (rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["EmailRecepcion"].ToString().Length != 0 && rows["EmailRecepcion"].ToString() != null &&
                                rows["PassEmailRecepcion"].ToString().Length != 0 && rows["PassEmailRecepcion"].ToString() != null &&
                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["SelectPrestamo"].ToString().Length != 0 && rows["SelectPrestamo"].ToString() != null &&
                                rows["ValorSol"].ToString().Length != 0 && rows["ValorSol"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["Contenido"].ToString().Length != 0 && rows["Contenido"].ToString() != null &&
                                rows["IngresosFam"].ToString().Length != 0 && rows["IngresosFam"].ToString() != null &&
                                rows["DedExterna"].ToString().Length != 0 && rows["DedExterna"].ToString() != null &&
                                rows["CompVigentes"].ToString().Length != 0 && rows["CompVigentes"].ToString() != null &&
                                rows["ValInmueble"].ToString().Length != 0 && rows["ValInmueble"].ToString() != null &&
                                rows["ValRemodel"].ToString().Length != 0 && rows["ValRemodel"].ToString() != null &&
                                rows["FechaSim1"].ToString().Length != 0 && rows["FechaSim1"].ToString() != null &&
                                rows["FechaSim2"].ToString().Length != 0 && rows["FechaSim2"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string CodPagina = rows["CodPagina"].ToString();
                                string EmailRecepcion = rows["EmailRecepcion"].ToString();
                                string PassEmailRecepcion = rows["PassEmailRecepcion"].ToString();
                                string Descripcion = rows["Descripcion"].ToString();
                                string SelectPrestamo = rows["SelectPrestamo"].ToString();
                                string ValorSol = rows["ValorSol"].ToString();

                                string user = rows["user"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string Contenido = rows["Contenido"].ToString();
                                string IngresosFam = rows["IngresosFam"].ToString();
                                string DedExterna = rows["DedExterna"].ToString();
                                string CompVigentes = rows["CompVigentes"].ToString();
                                string ValInmueble = rows["ValInmueble"].ToString();
                                string ValRemodel = rows["ValRemodel"].ToString();
                                string FechaSim1 = rows["FechaSim1"].ToString();
                                string FechaSim2 = rows["FechaSim2"].ToString();
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

                                    List<string> errorsTest = new List<string>();
                                    List<string> errors = new List<string>();
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];


                                    if (database == "SQL")
                                    {
                                        //ELIMINAR REGISTRO
                                        string eliminarRegistro2 = $"DELETE FROM bp_dsoli where act_usua='202020' and cod_empr='9'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro = $"DELETE FROM bp_sopre where cod_empl='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);

                                    }
                                    else
                                    {
                                        string eliminarRegistro2 = $"DELETE bp_dsoli where act_usua='193454' and cod_empr='421'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro = $"DELETE FROM bp_sopre where cod_empl='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                    }
                                    


                                    //INICIO PRUEBA
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//a[contains(.,'Mis Prestamos')]");
                                    selenium.Click("//a[contains(.,'Mis Prestamos')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Préstamos", true, file);

                                    Thread.Sleep(500);
                                    selenium.Click("//*[@id=\"ctl00_btnNuevo\"]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("MIS SOLICITUDES DE PRÉSTAMOS", true, file);
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//input[contains(@id,'txtIngFami')]");
                                    selenium.SendKeys("//input[contains(@id,'txtIngFami')]", IngresosFam);
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//input[contains(@id,'txtDedExte')]");
                                    selenium.SendKeys("//input[contains(@id,'txtDedExte')]", DedExterna);
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//input[contains(@id,'txtComVige')]");
                                    selenium.SendKeys("//input[contains(@id,'txtComVige')]", CompVigentes);
                                    selenium.Screenshot("Datos Ingresados", true, file);
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//input[contains(@id,'kcfFecSoliA_txtFecha')]");
                                    selenium.SendKeys("//input[contains(@id,'kcfFecSoliA_txtFecha')]", Fecha);
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//select[contains(@id,'ddlNomPres')]");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlNomPres')]", SelectPrestamo);
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//input[contains(@id,'txtValInmu')]");
                                    selenium.SendKeys("//input[contains(@id,'txtValInmu')]", ValInmueble);
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//input[contains(@id,'txtValRemo')]");
                                    selenium.SendKeys("//input[contains(@id,'txtValRemo')]", ValRemodel);
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtPlaPres']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtPlaPres']", "2");
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//input[contains(@id,'txtValPres')]");
                                    selenium.SendKeys("//input[contains(@id,'txtValPres')]", ValorSol);
                                    Thread.Sleep(1000);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Ingresados Formulario", true, file);
                                    //SIMULADOR
                                    selenium.Scroll("//a[contains(text(),'Simulador')]");
                                    selenium.Click("//a[contains(text(),'Simulador')]");
                                    Thread.Sleep(2000);
                                    //CUOTAS MENSUALES
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblCuotas_1']");
                                    Thread.Sleep(2000);
                                    //FECHAS
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_KCtrlFechaDesem_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFechaDesem_txtFecha']", FechaSim1);
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_KCtrlFechaDescue_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFechaDescue_txtFecha']", FechaSim2);
                                    Thread.Sleep(1000);
                                    //VERIFICAR
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnVerificar']");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Simulador Préstamo", true, file);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnSalir']");
                                    Thread.Sleep(1000);
                                    //Guardar
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                    {
                                        selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    }
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Datos Registrados", true, file);
                                    Thread.Sleep(4000);
                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(10000);
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
                                    if (errorsTest.Count > 0)
                                    {
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, errorsTest);

                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}",
                                                        Environment.NewLine, errorMessageString));
                                    }

                                    Thread.Sleep(2000);
                                    
                                    bandera = true;

                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();
                                        int StCount = Int32.Parse(SthCount);

                                        if (StCount > 0)
                                        {
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
        public void BP_SolicitudDePréstamos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_SolicitudDePréstamos")
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

                            if (rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&

                                rows["IngresosFami"].ToString().Length != 0 && rows["IngresosFami"].ToString() != null &&
                                rows["Externas"].ToString().Length != 0 && rows["Externas"].ToString() != null &&
                                rows["Vigentes"].ToString().Length != 0 && rows["Vigentes"].ToString() != null &&
                                rows["ValorInmueble"].ToString().Length != 0 && rows["ValorInmueble"].ToString() != null &&
                                rows["SelectPrestamo"].ToString().Length != 0 && rows["SelectPrestamo"].ToString() != null &&
                                rows["ValorSol"].ToString().Length != 0 && rows["ValorSol"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["Ruta"].ToString().Length != 0 && rows["Ruta"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["ValorRemodela"].ToString().Length != 0 && rows["ValorRemodela"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string IngresosFami = rows["IngresosFami"].ToString();
                                string Externas = rows["Externas"].ToString();
                                string Vigentes = rows["Vigentes"].ToString();
                                string ValorInmueble = rows["ValorInmueble"].ToString();
                                string SelectPrestamo = rows["SelectPrestamo"].ToString();
                                string ValorSol = rows["ValorSol"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string ValorRemodela = rows["ValorRemodela"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string FechaSim1 = rows["FechaSim1"].ToString();
                                string FechaSim2 = rows["FechaSim2"].ToString();

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
                                    List<string> errorsTest = new List<string>();
                                    List<string> errors = new List<string>();
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    if (database == "SQL")
                                    {
                                        //ELIMINAR REGISTRO
                                        string eliminarRegistro2 = $"DELETE FROM bp_dsoli where act_usua='4' and cod_empr='9'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro = $"DELETE FROM bp_sopre where cod_empl='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);

                                    }
                                    else
                                    {
                                        string eliminarRegistro2 = $"DELETE FROM bp_dsoli where act_usua='193454' and cod_empr='421'";
                                        db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                        string eliminarRegistro = $"DELETE FROM bp_sopre where cod_empl='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                    }

                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    //MIS PRESTAMOS
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//a[contains(.,'Mis Prestamos')]");
                                    selenium.Click("//a[contains(.,'Mis Prestamos')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Préstamos", true, file);
                                    //NUEVO
                                    Thread.Sleep(500);
                                    selenium.Click("//*[@id=\"ctl00_btnNuevo\"]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("MIS SOLICITUDES DE PRÉSTAMOS", true, file);
                                    Thread.Sleep(1000);
                                    //INGRESOS FAMILIARES
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtIngFami']", IngresosFami);
                                    //EXTERNAS
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtDedExte']", Externas);
                                    //VIGENTES
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtComVige']", Vigentes);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Datos Ingresados Formulario", true, file);
                                    //FECHAS
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_kcfFecSoliA_txtFecha']");
                                    Thread.Sleep(1000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_kcfFecSoliA_txtFecha']", Fecha);
                                    Thread.Sleep(1000);
                                    //PRESTAMOS
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlNomPres']");
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomPres']", SelectPrestamo);
                                    Thread.Sleep(1000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValInmu']", ValorInmueble);
                                    //VALOR INMUEBLE
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtValRemo']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValRemo']", ValorRemodela);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtPlaPres']", "2");
                                    Thread.Sleep(1000);
                                    //VALOR SOLICITUD
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtValPres']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtValPres']", ValorSol);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Datos Ingresados Formulario", true, file);
                                    //SIMULADOR
                                    selenium.Scroll("//a[contains(text(),'Simulador')]");
                                    selenium.Click("//a[contains(text(),'Simulador')]");
                                    Thread.Sleep(2000);

                                    //CUOTAS MENSUALES
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblCuotas_1']");
                                    Thread.Sleep(2000);

                                    //FECHAS
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_KCtrlFechaDesem_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFechaDesem_txtFecha']", FechaSim1);
                                    Thread.Sleep(1000);

                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_KCtrlFechaDescue_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFechaDescue_txtFecha']", FechaSim2);
                                    Thread.Sleep(1000);

                                    //VERIFICAR
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnVerificar']");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Simulador Prestamo", true, file);

                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnSalir']");
                                    Thread.Sleep(1000);

                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(3000);
                                    if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                    {
                                        selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    }
                                    Thread.Sleep(2000);
                                    //MIS PRESTAMOS
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//a[contains(.,'Mis Prestamos')]");
                                    selenium.Click("//a[contains(.,'Mis Prestamos')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgBpSopre_ctl03_LinkButton1']/i");
                                    selenium.Screenshot("Solicitud Registrada", true, file);
                                    //DETALLE
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgBpSopre_ctl03_LinkButton1']/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Ver Detalle", true, file);
                                    //ADJUNTAR ARCHIVO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtValPres']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtValPres']");
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait(Ruta);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    //TIPO DOCUMENTO
                                    if (database == "ORA")
                                    {
                                        selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_KCtrTipoDocumento_ddlTIP_DOCU']", "PRUEBAS");
                                        Thread.Sleep(1000);
                                    }
                                    else
                                    {
                                        selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_KCtrTipoDocumento_ddlTIP_DOCU']", "COPIA DE LA CEDULA");
                                        Thread.Sleep(1000);
                                    }
                                    //ARCHIVO ADJUNTO GUARDADO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnAdiArch']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Archivo Adjunto", true, file);
                                    //MIS PRESTAMOS
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//a[contains(.,'Mis Prestamos')]");
                                    selenium.Click("//a[contains(.,'Mis Prestamos')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Préstamos registrado", true, file);

                                    fv.ConvertWordToPDF(file, database); 
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

                                    Thread.Sleep(2000);
                                
                               
                                    bandera = true;

                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();
                                        int StCount = Int32.Parse(SthCount);

                                        if (StCount > 0)
                                        {
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
        public void BP_FlujoAprobaciónBeneficioDePersonalRolRRHH()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_FlujoAprobaciónBeneficioDePersonalRolRRHH")
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
                                rows["EmpleadoUser1"].ToString().Length != 0 && rows["EmpleadoUser1"].ToString() != null &&
                                rows["EmpleadoPass1"].ToString().Length != 0 && rows["EmpleadoPass1"].ToString() != null &&
                                rows["EmpleadoUser2"].ToString().Length != 0 && rows["EmpleadoUser2"].ToString() != null &&
                                rows["EmpleadoUser2"].ToString().Length != 0 && rows["EmpleadoUser2"].ToString() != null &&
                                rows["EmpleadoUser3"].ToString().Length != 0 && rows["EmpleadoUser3"].ToString() != null &&
                                rows["EmpleadoUser3"].ToString().Length != 0 && rows["EmpleadoUser3"].ToString() != null &&
                                rows["EmpleadoUser4"].ToString().Length != 0 && rows["EmpleadoUser4"].ToString() != null &&
                                rows["EmpleadoUser4"].ToString().Length != 0 && rows["EmpleadoUser4"].ToString() != null &&
                                //Variables
                                rows["Empresa"].ToString().Length != 0 && rows["Empresa"].ToString() != null &&
                                rows["Valor"].ToString().Length != 0 && rows["Valor"].ToString() != null &&
                                rows["Beneficio"].ToString().Length != 0 && rows["Beneficio"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string EmpleadoUser1 = rows["EmpleadoUser1"].ToString();
                                string EmpleadoPass1 = rows["EmpleadoPass1"].ToString();
                                string EmpleadoUser2 = rows["EmpleadoUser2"].ToString();
                                string EmpleadoPass2 = rows["EmpleadoPass2"].ToString();
                                string EmpleadoUser3 = rows["EmpleadoUser3"].ToString();
                                string EmpleadoPass3 = rows["EmpleadoPass3"].ToString();
                                string EmpleadoUser4 = rows["EmpleadoUser4"].ToString();
                                string EmpleadoPass4 = rows["EmpleadoPass4"].ToString();
                                string user = rows["user"].ToString();
                                //Variables
                                string Empresa = rows["Empresa"].ToString();
                                string Beneficio = rows["Beneficio"].ToString();
                                string Valor = rows["Valor"].ToString();
                                string Modalidad = rows["Modalidad"].ToString();
                                string Programa = rows["Programa"].ToString();
                                string Calendario = rows["Calendario"].ToString();
                                string TipIntensidad = rows["TipIntensidad"].ToString();
                                string TipDocumento = rows["TipDocumento"].ToString();

                                string ValorSolicitado = rows["ValorSolicitado"].ToString();
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
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    if (database == "SQL")
                                    {
                                        string EliminarConsecutivo = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo, database, user);

                                        string EliminarDetalle = $"Delete from BP_OTOBE where COD_BENE ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarDetalle, database, user);

                                        string EliminarConsecutivo1 = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{EmpleadoUser1}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo1, database, user);

                                    }
                                    else
                                    {
                                        string EliminarConsecutivo = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo, database, user);

                                        string EliminarDetalle = $"Delete from BP_OTOBE where COD_BENE ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(EliminarDetalle, database, user);

                                        string EliminarConsecutivo1 = $"Delete from NM_SOLTR where tip_apli ='B' AND COD_RESP ='{EmpleadoUser1}'";
                                        db.UpdateDeleteInsert(EliminarConsecutivo1, database, user);

                                    }
                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//a[contains(.,'Beneficios Organizacionales')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//a[contains(.,'Beneficios Organizacionales')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Solicitud de Beneficios", true, file);
                                    if (database == "ORA")
                                    {
                                        selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_rbBenef_0')]");
                                        Thread.Sleep(1500);
                                        selenium.Scroll("//select[contains(@id,'ctl00_ContenidoPagina_ddlNomTibe')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlNomTibe')]", Beneficio);
                                        Thread.Sleep(1500);
                                        selenium.Scroll("//select[contains(@id,'ctl00_ContenidoPagina_ddlEntBen')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlEntBen')]", "Prueba");
                                        Thread.Sleep(1500);
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_txtValSoli')]");
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtValSoli')]", Valor);
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Datos de la Solicitud", true, file);
                                        //AGREGAR OBSERVACIONES
                                        selenium.Scroll("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]");
                                        selenium.SendKeys("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]", Observaciones);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Observaciones", true, file);
                                        //ADJUNTO
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(5000);
                                        
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(2000);
                                        SendKeys.SendWait(@"C:\Reportes\ArchivoSelf\ArchivoPrueba.pdf");
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(2000);
                                    }
                                    else
                                    {
                                        //BENEFICIOS

                                        selenium.Click("//input[contains(@id,'rbBenef_0')]");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Beneficio colaborador", true, file);
                                        //TIPO DE AUXILIO
                                        selenium.SelectElementByName("//select[contains(@id,'ddlNomTibe')]", "AUXILIO DE ANTEOJOS");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Tipo auxilio", true, file);

                                        //INGRESAR VALOR SOLICITADO
                                        selenium.Scroll("//input[contains(@id,'txtValSoli')]");
                                        Thread.Sleep(2000);
                                        selenium.SendKeys("//input[contains(@id,'txtValSoli')]", ValorSolicitado);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Valor Solicitado", true, file);
                                        SendKeys.SendWait("{TAB}");
                                        //AGREGAR OBSERVACIONES
                                        selenium.Scroll("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]");
                                        selenium.SendKeys("//textarea[contains(@id,'KtxtObserSoli_txtTexto')]", Observaciones);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Observaciones", true, file);
                                        //ADJUNTO
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(5000);
                                        
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(2000);
                                        SendKeys.SendWait(@"C:\Reportes\ArchivoSelf\ArchivoPrueba.pdf");
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(2000);
                                    }
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Solicitud Radicada", true, file);
                                    if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                    {
                                        selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    }
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Verificar Solicitud", true, file);
                                    selenium.Close();

                                    //Aprobador 1
                                    selenium.LoginApps(app, EmpleadoUser1, EmpleadoPass1, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Aprobador RRHH", true, file);
                                    Thread.Sleep(1000);

                                    if (database == "SQL")
                                    {
                                        selenium.Click("//button[contains(.,'Rol RRHH')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//*[@id='ctl00_pRRHH']");
                                    }
                                    selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Rol RRHH", true, file);
                                    Thread.Sleep(1000);
                                   
                                    selenium.Scroll("//a[contains(.,'BENEFICIOS ORGANIZACIONALES')]");
                                    selenium.Click("//a[contains(.,'BENEFICIOS ORGANIZACIONALES')]");
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//a[contains(@href, 'frmRHBpBeotoL.aspx')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(@href, 'frmRHBpBeotoL.aspx')]");
                                    Thread.Sleep(4000);
                                    //filtro fecha
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']");
                                    Thread.Sleep(2000);
                                    string date = DateTime.UtcNow.ToString("dd/MM/yyyy");
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", date);
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']", date);
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_btnConsultar']");
                                    Thread.Sleep(2000);
                                    //filtro empleado
                                    selenium.Click("//*[@id='tblBpBeoto_filter']/label/input");
                                    selenium.SendKeys("//*[@id='tblBpBeoto_filter']/label/input", "AUXILIO");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle", true, file);
                                    
                                    selenium.Click("//*[@id='tblBpBeoto']/tbody/tr/td[8]/a");
                                    Thread.Sleep(1000);
                                   
                                    //APROBAR
                                    selenium.Screenshot("Aprobación", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_Aprueba')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Enviar Correo Aprobación", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnEnviar')]");
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    //VALIDACION SOLICITUD APROBADA
                                    //Login Empleado
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login Empleado", true, file);
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//a[contains(.,'MIS SOLICITUDES')]");
                                    Thread.Sleep(1000);
                                    selenium.Scroll("(//a[contains(@href, 'frmBpBeotoL.aspx')])[2]");
                                    Thread.Sleep(1000);
                                    selenium.Click("(//a[contains(@href, 'frmBpBeotoL.aspx')])[2]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Solicitud de Beneficios", true, file);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(3000);
                                    //DETALLE
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgBpBeoto_ctl03_LinkButton1']/i");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dgHistObser_ctl03_LinkButton4']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Aprobacion RRHH", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    ////////////////////////////////////////////////////
                                    ///////////////////////////////////////////////////////
                                    //limpiar procesos
                                    Process[] processes1 = Process.GetProcessesByName("chromedriver");
                                    if (processes1.Length > 0)
                                    {
                                        for (int i = 0; i < processes1.Length; i++)
                                        {
                                            processes1[i].Kill();
                                        }
                                    }
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
        public void BP_AsignaciónInsigniasEquipoaCargoD()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_AsignaciónInsigniasEquipoaCargoD")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Datos Prueba    
                                rows["Persona"].ToString().Length != 0 && rows["Persona"].ToString() != null &&
                                rows["Mensaje"].ToString().Length != 0 && rows["Mensaje"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Persona = rows["Persona"].ToString();
                                string Mensaje = rows["Mensaje"].ToString();
                                string url2 = rows["url2"].ToString();
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
                                    if (database == "SQL")
                                    {

                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 507195 AND COD_EMPR = 9";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

                                        string actualizar2 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 79753160 AND COD_EMPR = 9";
                                        db.UpdateDeleteInsert(actualizar2, database, user);

                                        string actualizar3 = $"update nm_contr set cod_frep = 405, ind_Acti = 'A' where cod_empl IN (507195) and cod_empr = 9";
                                        db.UpdateDeleteInsert(actualizar3, database, user);

                                        string actualizar4 = $"update nm_contr set cod_ccos = 1 where cod_empl IN (507195 ) and cod_empr = 9 and ind_Acti = 'A'";
                                        db.UpdateDeleteInsert(actualizar4, database, user);

                                        string actualizar5 = $"update nm_contr set cod_ccos = 2 where cod_empl IN (79753160 ) and cod_empr = 9 and ind_Acti = 'A'";
                                        db.UpdateDeleteInsert(actualizar5, database, user);

                                        string actualizar6 = $"update nm_contr set cod_frep = 2020 where cod_empl IN (79753160) and cod_empr = 9";
                                        db.UpdateDeleteInsert(actualizar6, database, user);

                                        string actualizar7 = $"update bi_cargo set cod_nive = 1 where cod_carg IN ('1','5') and cod_empr = 9";
                                        db.UpdateDeleteInsert(actualizar7, database, user);

                                    }
                                    else
                                    {
                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

                                        string actualizar2 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 94541552 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar2, database, user);

                                        string actualizar3 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 22548965 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar3, database, user);

                                        string actualizar4 = $"update nm_contr set cod_frep = 94541552 where cod_empl in (19301797)";
                                        db.UpdateDeleteInsert(actualizar4, database, user);

                                        string actualizar5 = $"update nm_contr set cod_frep = 84082016 where cod_empl in (22548965)";
                                        db.UpdateDeleteInsert(actualizar5, database, user);

                                        string actualizar6 = $"update nm_contr set cod_ccos = 1 where cod_empl IN (19301797,22548965) and cod_empr = 421 and ind_Acti = 'A'";
                                        db.UpdateDeleteInsert(actualizar6, database, user);

                                        string actualizar7 = $"update nm_contr set cod_ccos = 1000 where cod_empl IN (22548965) and cod_empr = 421 and ind_Acti = 'A'";
                                        db.UpdateDeleteInsert(actualizar7, database, user);

                                        string actualizar8 = $"update bi_cargo set cod_nive = 3 where cod_carg IN ('204411','ABC123') and cod_empr = 421";
                                        db.UpdateDeleteInsert(actualizar8, database, user);

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

                                    //INGRESO A INSIGNIAS
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INSIGNIAS", true, file);

                                    //SELECCION PERSONA
                                    selenium.Click("//span[@id='select2-ctl00_ContenidoPagina_ddlColaborador-container']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@type='search']", Persona);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("PERSONA", true, file);
                                    selenium.Enter("//input[@type='search']");

                                    //MENSAJE
                                    selenium.Click("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']", Mensaje);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MENSAJE", true, file);

                                    selenium.Scroll("//div[@id='1']/div");

                                    //ARRASTRAR INSIGNIA

                                    IWebElement desde = driver2.FindElement(By.XPath("//div[@id='1']/div"));
                                    IWebElement hasta = driver2.FindElement(By.XPath("//div[@id='container-right']"));

                                    Actions action = new Actions(driver2);
                                    action.DragAndDrop(desde, hasta).Perform();

                                    selenium.Screenshot("INSIGNIA ARRASTRADA", true, file);
                                    Thread.Sleep(2000);

                                    //AUTODIDACTA
                                    selenium.Click("//div[@id='divComportamientos']/div/div/label/span");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("AUTODIDACTA", true, file);

                                    //CONFIRMAR AUTODIDACTA
                                    selenium.Scroll("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(2000);

                                    //CONFIRMAR
                                    selenium.Scroll("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Click("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("SATISFACTORIO", true, file);

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
        public void BP_CentroVacacionesTérminosCondiciones()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_CentroVacacionesTérminosCondiciones")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba    
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFin"].ToString().Length != 0 && rows["FechaFin"].ToString() != null &&
                                rows["FormaPago"].ToString().Length != 0 && rows["FormaPago"].ToString() != null &&
                                rows["FamiliarID"].ToString().Length != 0 && rows["FamiliarID"].ToString() != null &&
                                rows["FamiliarNombre"].ToString().Length != 0 && rows["FamiliarNombre"].ToString() != null &&
                                rows["FamiliarApellido"].ToString().Length != 0 && rows["FamiliarApellido"].ToString() != null &&
                                rows["FamiliarFecha"].ToString().Length != 0 && rows["FamiliarFecha"].ToString() != null &&
                                rows["Prestamo"].ToString().Length != 0 && rows["Prestamo"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFin = rows["FechaFin"].ToString();
                                string FormaPago = rows["FormaPago"].ToString();
                                string FamiliarID = rows["FamiliarID"].ToString();
                                string FamiliarNombre = rows["FamiliarNombre"].ToString();
                                string FamiliarApellido = rows["FamiliarApellido"].ToString();
                                string FamiliarFecha = rows["FamiliarFecha"].ToString();
                                string Prestamo = rows["Prestamo"].ToString();
                                string JefeUser1 = rows["JefeUser1"].ToString();
                                string JefeUser2 = rows["JefeUser2"].ToString();
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

                                    
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               
                                    if (database == "SQL")
                                    {
                                        string familiar = $"delete from BP_DSCEV where ACT_USUA='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(familiar, database, user);

                                        string centro = $"delete from bp_cenva where COD_EMPR= 9 AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(centro, database, user);

                                        string empleado = $"delete from nm_soltr where COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(empleado, database, user);

                                        string jefe1 = $"delete from nm_soltr where COD_RESP = '{JefeUser1}' AND TIP_APLI='Q'";
                                        db.UpdateDeleteInsert(jefe1, database, user);

                                        string jefe2 = $"delete from nm_soltr where COD_RESP = '{JefeUser2}' AND TIP_APLI='Q'";
                                        db.UpdateDeleteInsert(jefe2, database, user);
                                    }
                                    else
                                    {
                                        string familiar = $"delete from BP_DSCEV where ACT_USUA='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(familiar, database, user);

                                        string centro = $"delete from bp_cenva where COD_EMPR= 421 AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(centro, database, user);

                                        string empleado = $"delete from nm_soltr where COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(empleado, database, user);

                                        string jefe1 = $"delete from nm_soltr where COD_RESP = '{JefeUser1}' AND TIP_APLI='Q'";
                                        db.UpdateDeleteInsert(jefe1, database, user);

                                        string jefe2 = $"delete from nm_soltr where COD_RESP = '{JefeUser2}' AND TIP_APLI='Q'";
                                        db.UpdateDeleteInsert(jefe2, database, user);
                                    }
                                    

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A CENTRO VACACIONAL
                                    selenium.Scroll("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);

                                    if (database == "ORA")
                                    {
                                        selenium.Scroll("//*[@id='MenuContex']/div[2]/div[1]/ul[1]/li[22]/ul[1]/li[1]/a[1]");
                                        selenium.Click("//*[@id='MenuContex']/div[2]/div[1]/ul[1]/li[22]/ul[1]/li[1]/a[1]");
                                        Thread.Sleep(2000);
                                    }
                                    else
                                    {
                                        selenium.Scroll("//a[contains(@href, 'frmBpCenvacaL.aspx')]");
                                        selenium.Click("//a[contains(@href, 'frmBpCenvacaL.aspx')]");
                                        Thread.Sleep(2000);
                                    }
                                    selenium.Screenshot("MIS CENTROS VACACIONALES", true, file);
                                    //SELECCION CENTRO
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//a/div/div/img");
                                        Thread.Sleep(10000);
                                        selenium.Screenshot("Solicitudes Centros Vacacionales", true, file);
                                    }
                                    else
                                    {
                                        selenium.Click("//a/div/div/img");
                                        Thread.Sleep(10000);
                                        selenium.Screenshot("Solicitudes Centros Vacacionales", true, file);
                                    }

                                    //FECHAS
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFechIni_txtFecha']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFechIni_txtFecha']", FechaInicial);
                                    Thread.Sleep(10000);
                                    selenium.Screenshot("Fecha Inicial", true, file);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFechFin_txtFecha']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFechFin_txtFecha']", FechaFin);
                                    selenium.Screenshot("Fecha Final", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(10000);
                                    //CHECK HACE PARTE GRUPO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_rblUstHues_0']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblUstHues_0']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Parte Acompañantes", true, file);
                                    //FAMILIAR
                                    selenium.Scroll("//div[@id='ctl00_ContenidoPagina_ddlFamil_sl']/div");
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_ddlFamil_sl']/div");
                                    Thread.Sleep(3000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_ddlFamil_0']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Seleccionar Familiar", true, file);
                                    selenium.Click("//input[@value='Aceptar']");
                                    Thread.Sleep(3000);
                                    //FORMA PAGO
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_rdbforpago']");
                                    Thread.Sleep(3000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_rdbforpago']", FormaPago);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Forma Pago Préstamo", true, file);
                                    //TOTAL A PAGAR
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_btnValCan']/span");
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_btnValCan']/span");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Total Pagar", true, file);
                                    //OBSERVACIONES
                                    selenium.Scroll("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValObser_txtTexto']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValObser_txtTexto']", "PRUEBAS CENTROS VACAIONALES");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Observaciones", true, file);

                                    //ADICIONAR ACOMPAÑANTES
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_OtroAcomp']");
                                    Thread.Sleep(3000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_OtroAcomp']");
                                    Thread.Sleep(3000);
                                    //ID ACOMPAÑANTE
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtidAcomp']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtidAcomp']", FamiliarID);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("ID Acompañante", true, file);
                                    //NOMBRE ACOMPAÑANTE
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtNombre']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNombre']", FamiliarNombre);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nombre Acompañante", true, file);
                                    //APELLIDO ACOMPAÑANTE
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtApellido']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtApellido']", FamiliarApellido);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Apellido Acompañante", true, file);
                                    //FECHA ACOMPAÑANTE
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_KCtrlFecAcomp_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecAcomp_txtFecha']", FamiliarFecha);
                                    Thread.Sleep(3000);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Screenshot("Fecha Nacimiento Acompañante", true, file);
                                    //GUARDAR ACOMPAÑANTE
                                    selenium.Scroll("//a[contains(text(),'Guardar acompañante')]");
                                    selenium.Click("//a[contains(text(),'Guardar acompañante')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Acompañante Agregado", true, file);
                                    //GUARDAR
                                    selenium.Scroll("//a[contains(text(),'Aceptar')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Acompañante Agregado", true, file);
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[contains(text(),'Aceptar')]");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Terminos y Condiciones", true, file);
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
        public void BP_AsignaciónInsigniasEquipoaCargoC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_AsignaciónInsigniasEquipoaCargoC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Datos Prueba    
                                rows["Persona"].ToString().Length != 0 && rows["Persona"].ToString() != null &&
                                rows["Mensaje"].ToString().Length != 0 && rows["Mensaje"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Persona = rows["Persona"].ToString();
                                string Mensaje = rows["Mensaje"].ToString();
                                string url2 = rows["url2"].ToString();
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
                                    if (database == "SQL")
                                    {

                                        string actualizar = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 507195 AND COD_EMPR = 9";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 79753160 AND COD_EMPR = 9";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

                                        string actualizar2 = $"update nm_contr set cod_frep = 405, ind_Acti = 'A' where cod_empl IN (507195) and cod_empr = 9";
                                        db.UpdateDeleteInsert(actualizar2, database, user);

                                        string actualizar3 = $"update nm_contr set cod_ccos = 1 where cod_empl IN (507195,79753160 ) and cod_empr = 9 and ind_Acti = 'A'";
                                        db.UpdateDeleteInsert(actualizar3, database, user);

                                        string actualizar4 = $"update nm_contr set cod_frep = 124 where cod_empl IN (79753160) and cod_empr = 9";
                                        db.UpdateDeleteInsert(actualizar4, database, user);

                                        string actualizar5 = $"update bi_cargo set cod_nive = 1 where cod_carg IN ('1','5') and cod_empr = 9";
                                        db.UpdateDeleteInsert(actualizar5, database, user);

                                    }
                                    else
                                    {
                                        string actualizar = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 94541552 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

                                        string actualizar2 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 22548965 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar2, database, user);

                                        string actualizar3 = $"update nm_contr set cod_frep = 94541552 where cod_empl in (19301797)";
                                        db.UpdateDeleteInsert(actualizar3, database, user);

                                        string actualizar4 = $"update nm_contr set cod_frep = 84082016 where cod_empl in (22548965)";
                                        db.UpdateDeleteInsert(actualizar4, database, user);

                                        string actualizar5 = $"update nm_contr set cod_ccos = 1 where cod_empl IN (19301797,22548965) and cod_empr = 421 and ind_Acti = 'A'";
                                        db.UpdateDeleteInsert(actualizar5, database, user);

                                        string actualizar6 = $"update bi_cargo set cod_nive = 3 where cod_carg IN ('204411','ABC123') and cod_empr = 421";
                                        db.UpdateDeleteInsert(actualizar6, database, user);
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

                                    //INGRESO A INSIGNIAS
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INSIGNIAS", true, file);

                                    //SELECCION PERSONA
                                    selenium.Click("//span[@id='select2-ctl00_ContenidoPagina_ddlColaborador-container']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@type='search']", Persona);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("PERSONA", true, file);
                                    selenium.Enter("//input[@type='search']");

                                    //MENSAJE
                                    selenium.Click("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']", Mensaje);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MENSAJE", true, file);

                                    selenium.Scroll("//div[@id='1']/div");

                                    //ARRASTRAR INSIGNIA

                                    IWebElement desde = driver2.FindElement(By.XPath("//div[@id='1']/div"));
                                    IWebElement hasta = driver2.FindElement(By.XPath("//div[@id='container-right']"));

                                    Actions action = new Actions(driver2);
                                    action.DragAndDrop(desde, hasta).Perform();

                                    selenium.Screenshot("INSIGNIA ARRASTRADA", true, file);
                                    Thread.Sleep(2000);

                                    //AUTODIDACTA
                                    selenium.Click("//div[@id='divComportamientos']/div/div/label/span");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("AUTODIDACTA", true, file);

                                    //CONFIRMAR AUTODIDACTA
                                    selenium.Scroll("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(2000);

                                    //CONFIRMAR
                                    selenium.Scroll("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Click("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("SATISFACTORIO", true, file);

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
        public void BP_AsignaciónInsigniasEquipoaCargoB()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_AsignaciónInsigniasEquipoaCargoB")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Datos Prueba    
                                rows["Persona"].ToString().Length != 0 && rows["Persona"].ToString() != null &&
                                rows["Mensaje"].ToString().Length != 0 && rows["Mensaje"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Persona = rows["Persona"].ToString();
                                string Mensaje = rows["Mensaje"].ToString();
                                string url2 = rows["url2"].ToString();
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
                                    if (database == "SQL")
                                    {

                                        string actualizar = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 507195 AND COD_EMPR = 9";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 79753160 AND COD_EMPR = 9";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

                                        string actualizar2 = $"update nm_contr set cod_frep = 405, ind_Acti = 'A' where cod_empl IN (507195,79753160 ) and cod_empr = 9";
                                        db.UpdateDeleteInsert(actualizar2, database, user);



                                    }
                                    else
                                    {
                                        string actualizar = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 94541552 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

                                        string actualizar2 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 22548965 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar2, database, user);

                                        string actualizar3 = $"update nm_contr set cod_frep = 94541552 where cod_empl in (19301797,22548965)";
                                        db.UpdateDeleteInsert(actualizar3, database, user);


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

                                    //INGRESO A INSIGNIAS
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INSIGNIAS", true, file);

                                    //SELECCION PERSONA
                                    selenium.Click("//span[@id='select2-ctl00_ContenidoPagina_ddlColaborador-container']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@type='search']", Persona);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("PERSONA", true, file);
                                    selenium.Enter("//input[@type='search']");

                                    //MENSAJE
                                    selenium.Click("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']", Mensaje);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MENSAJE", true, file);

                                    selenium.Scroll("//div[@id='1']/div");

                                    //ARRASTRAR INSIGNIA

                                    IWebElement desde = driver2.FindElement(By.XPath("//div[@id='1']/div"));
                                    IWebElement hasta = driver2.FindElement(By.XPath("//div[@id='container-right']"));

                                    Actions action = new Actions(driver2);
                                    action.DragAndDrop(desde, hasta).Perform();

                                    selenium.Screenshot("INSIGNIA ARRASTRADA", true, file);
                                    Thread.Sleep(2000);

                                    //AUTODIDACTA
                                    selenium.Click("//div[@id='divComportamientos']/div/div/label/span");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("AUTODIDACTA", true, file);

                                    //CONFIRMAR AUTODIDACTA
                                    selenium.Scroll("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(2000);

                                    //CONFIRMAR
                                    selenium.Scroll("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Click("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("SATISFACTORIO", true, file);

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
        public void BP_AsignaciónInsigniasEquipoaCargoA()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_AsignaciónInsigniasEquipoaCargoA")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Datos Prueba    
                                rows["Persona"].ToString().Length != 0 && rows["Persona"].ToString() != null &&
                                rows["Mensaje"].ToString().Length != 0 && rows["Mensaje"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Persona = rows["Persona"].ToString();
                                string Mensaje = rows["Mensaje"].ToString();
                                string url2 = rows["url2"].ToString();

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
                                    if (database == "SQL")
                                    {

                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                    }
                                    else
                                    {
                                        string actualizar = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
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

                                    //INGRESO A INSIGNIAS
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INSIGNIAS", true, file);

                                    //SELECCION PERSONA
                                    selenium.Click("//span[@id='select2-ctl00_ContenidoPagina_ddlColaborador-container']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@type='search']", Persona);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("PERSONA", true, file);
                                    SendKeys.SendWait("{DOWN}");
                                    selenium.Enter("//input[@type='search']");

                                    //MENSAJE
                                    selenium.Click("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']", Mensaje);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MENSAJE", true, file);

                                    selenium.Scroll("//div[@id='1']/div");

                                    //ARRASTRAR INSIGNIA

                                    IWebElement desde = driver2.FindElement(By.XPath("//div[@id='1']/div"));
                                    IWebElement hasta = driver2.FindElement(By.XPath("//div[@id='container-right']"));

                                    Actions action = new Actions(driver2);
                                    action.DragAndDrop(desde, hasta).Perform();

                                    selenium.Screenshot("INSIGNIA ARRASTRADA", true, file);
                                    Thread.Sleep(2000);

                                    //AUTODIDACTA
                                    selenium.Click("//div[@id='divComportamientos']/div/div/label/span");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("AUTODIDACTA", true, file);

                                    //CONFIRMAR AUTODIDACTA
                                    selenium.Scroll("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(2000);

                                    //CONFIRMAR
                                    selenium.Scroll("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Click("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("SATISFACTORIO", true, file);

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(5000);
                                    selenium.Close();
                                    string verA = string.Empty;
                                    string verB = string.Empty;
                                    //CONSULTA BASE DATOS
                                    if (database == "SQL")
                                    {
                                        string consulta = $"Select pun_roin from bp_asins where cod_empl = 507195 and OBS_ERVA LIKE 'CASO DE USO'  AND COD_ROLI = 'A'";
                                        DataTable resultado = db.Select(consulta, user, database);
                                        if (resultado.Rows.Count > 0)
                                        {
                                            foreach (DataRow rw in resultado.Rows)
                                            {
                                                verA = rw["PUN_ROIN"].ToString();

                                            }
                                        }
                                        //CONSULTA SQL
                                        string consulta1 = $"SELECT PUN_ROIN FROM bp_ropui where cod_roli= 'A' and cod_empr = 9";
                                        DataTable resultado1 = db.Select(consulta1, user, database);
                                        if (resultado1.Rows.Count > 0)
                                        {
                                            foreach (DataRow rw in resultado1.Rows)
                                            {
                                                verB = rw["PUN_ROIN"].ToString();

                                            }
                                        }

                                        APIFuncionesVitales.InsertConsulta(file, verA, consulta);
                                        APIFuncionesVitales.InsertConsulta(file, verB, consulta1);
                                    }
                                    else
                                    {
                                        string consulta = $"Select pun_roin from bp_asins where cod_empl = 19301797 and OBS_ERVA LIKE 'CASO DE USO'  AND COD_ROLI = 'A'";
                                        DataTable resultado = db.Select(consulta, user, database);
                                        if (resultado.Rows.Count > 0)
                                        {
                                            foreach (DataRow rw in resultado.Rows)
                                            {
                                                verA = rw["PUN_ROIN"].ToString();

                                            }
                                        }
                                        //CONSULTA SQL
                                        string consulta1 = $"SELECT PUN_ROIN FROM bp_ropui where cod_roli= 'A' and cod_empr = 421";
                                        DataTable resultado1 = db.Select(consulta1, user, database);
                                        if (resultado1.Rows.Count > 0)
                                        {
                                            foreach (DataRow rw in resultado1.Rows)
                                            {
                                                verB = rw["PUN_ROIN"].ToString();

                                            }
                                        }

                                        APIFuncionesVitales.InsertConsulta(file, verA, consulta);
                                        APIFuncionesVitales.InsertConsulta(file, verB, consulta1);
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
        public void BP_AsignaciónInsigniasLíderInmediatoE()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_AsignaciónInsigniasLíderInmediatoE")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Datos Prueba    
                                rows["Persona"].ToString().Length != 0 && rows["Persona"].ToString() != null &&
                                rows["Mensaje"].ToString().Length != 0 && rows["Mensaje"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Persona = rows["Persona"].ToString();
                                string Mensaje = rows["Mensaje"].ToString();
                                string url2 = rows["url2"].ToString();
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
                                    if (database == "SQL")
                                    {

                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);


                                    }
                                    else
                                    {
                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
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

                                    //INGRESO A INSIGNIAS
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INSIGNIAS", true, file);

                                    //SELECCION PERSONA
                                    selenium.Click("//span[@id='select2-ctl00_ContenidoPagina_ddlColaborador-container']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@type='search']", Persona);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("PERSONA", true, file);
                                    selenium.Enter("//input[@type='search']");

                                    //MENSAJE
                                    selenium.Click("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']", Mensaje);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MENSAJE", true, file);

                                    selenium.Scroll("//div[@id='1']/div");

                                    //ARRASTRAR INSIGNIA

                                    IWebElement desde = driver2.FindElement(By.XPath("//div[@id='1']/div"));
                                    IWebElement hasta = driver2.FindElement(By.XPath("//div[@id='container-right']"));

                                    Actions action = new Actions(driver2);
                                    action.DragAndDrop(desde, hasta).Perform();

                                    selenium.Screenshot("INSIGNIA ARRASTRADA", true, file);
                                    Thread.Sleep(2000);

                                    //AUTODIDACTA
                                    selenium.Click("//div[@id='divComportamientos']/div/div/label/span");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("AUTODIDACTA", true, file);

                                    //CONFIRMAR AUTODIDACTA
                                    selenium.Scroll("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(4000);
                                    selenium.Click("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(4000);

                                    //CONFIRMAR
                                    selenium.Scroll("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Click("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("SATISFACTORIO", true, file);

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(5000);
                                    selenium.Close();
                                    string verA = string.Empty;
                                    string verB = string.Empty;
                                    //CONSULTA SQL
                                    if (database == "SQL")
                                    {
                                        string consulta = $"Select pun_roin from bp_asins   where   cod_empl = 13  and OBS_ERVA LIKE 'CASO DE USO' ";
                                        DataTable resultado = db.Select(consulta, user, database);
                                        if (resultado.Rows.Count > 0)
                                        {
                                            foreach (DataRow rw in resultado.Rows)
                                            {
                                                verA = rw["PUN_ROIN"].ToString();

                                            }
                                        }
                                        //CONSULTA SQL
                                        string consulta1 = $"SELECT PUN_ROIN FROM bp_ropui where   cod_roli= 'E' and cod_empr = 9 ";
                                        DataTable resultado1 = db.Select(consulta1, user, database);
                                        if (resultado1.Rows.Count > 0)
                                        {
                                            foreach (DataRow rw in resultado1.Rows)
                                            {
                                                verB = rw["PUN_ROIN"].ToString();

                                            }
                                        }

                                        APIFuncionesVitales.InsertConsulta(file, verA, consulta);
                                        APIFuncionesVitales.InsertConsulta(file, verB, consulta1);

                                    }
                                    else
                                    {
                                        string consulta = $"Select pun_roin from bp_asins   where   cod_empl = 39801386  and OBS_ERVA LIKE 'CASO DE USO' ";
                                        DataTable resultado = db.Select(consulta, user, database);
                                        if (resultado.Rows.Count > 0)
                                        {
                                            foreach (DataRow rw in resultado.Rows)
                                            {
                                                verA = rw["PUN_ROIN"].ToString();

                                            }
                                        }
                                        //CONSULTA SQL
                                        string consulta1 = $"SELECT PUN_ROIN FROM bp_ropui where   cod_roli= 'E' and cod_empr = 421 ";
                                        DataTable resultado1 = db.Select(consulta1, user, database);
                                        if (resultado1.Rows.Count > 0)
                                        {
                                            foreach (DataRow rw in resultado1.Rows)
                                            {
                                                verB = rw["PUN_ROIN"].ToString();

                                            }
                                        }

                                        APIFuncionesVitales.InsertConsulta(file, verA, consulta);
                                        APIFuncionesVitales.InsertConsulta(file, verB, consulta1);

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
        public void BP_ValidacióninsigniasMiEquipoPorInsignia()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_ValidacióninsigniasMiEquipoPorInsignia")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&

                                //Datos Prueba
                                rows["Persona"].ToString().Length != 0 && rows["Persona"].ToString() != null



                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Persona = rows["Persona"].ToString();
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

                                        string actualizar = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update nm_contr set cod_frep = 19301797 where cod_empl = 19301797";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

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

                                    //INGRESO A MIS COLABORADORES/INSIGNIAS COLABORADORES
                                    selenium.Click("//a[contains(.,'MIS COLABORADORES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Insignias de mis colaboradores')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Insignias de mis colaboradores')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A INSIGNIAS DE MIS COLABORADORES", true, file);

                                    //BUSCAR
                                    selenium.Click("//div[@id='tablaInsiEquipo_filter']/label/input");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//div[@id='tablaInsiEquipo_filter']/label/input", Persona);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("BUSCAR PERSONA", true, file);

                                    //VER INISGNIAS

                                    selenium.Click("//button[contains(.,'Ver insignias adquiridas')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("ISNIGNIA", true, file);


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
        public void BP_ValidacióninsigniasMiEquipo()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_ValidacióninsigniasMiEquipo")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&

                                //Datos Prueba
                                rows["Persona"].ToString().Length != 0 && rows["Persona"].ToString() != null



                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Persona = rows["Persona"].ToString();
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

                                        string actualizar = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update nm_contr set cod_frep = 19301797 where cod_empl = 19301797";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

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

                                    //INGRESO A MIS COLABORADORES/INSIGNIAS COLABORADORES
                                    selenium.Click("//a[contains(.,'MIS COLABORADORES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Insignias de mis colaboradores')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Insignias de mis colaboradores')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A INSIGNIAS DE MIS COLABORADORES", true, file);

                                    //BUSCAR
                                    selenium.Click("//div[@id='tablaInsiEquipo_filter']/label/input");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//div[@id='tablaInsiEquipo_filter']/label/input", Persona);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("BUSCAR PERSONA", true, file);

                                    //VER INISGNIAS
                                    selenium.Screenshot("ISNIGNIA", true, file);


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
        public void BP_ValidaciónMiEquipoEnInsignias()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_ValidaciónMiEquipoEnInsignias")
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

                                        string actualizar = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update nm_contr set cod_frep = 19301797 where cod_empl = 19301797";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

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
                                    //INGRESO A MIS COLABORADORES/INSIGNIAS COLABORADORES
                                    selenium.Click("//a[contains(.,'MIS COLABORADORES')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Insignias de mis colaboradores')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Insignias de mis colaboradores')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A INSIGNIAS DE MIS COLABORADORES", true, file);

                                    //REGISTROS
                                    selenium.ScrollTo("0", "800");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("CANTIDAD REGISTROS", true, file);


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
        public void BP_EnvíoCorreoAsignacionInsignias()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_EnvíoCorreoAsignacionInsignias")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Datos Prueba    
                                rows["Persona"].ToString().Length != 0 && rows["Persona"].ToString() != null &&
                                rows["Mensaje"].ToString().Length != 0 && rows["Mensaje"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Persona = rows["Persona"].ToString();
                                string Mensaje = rows["Mensaje"].ToString();
                                string url2 = rows["url2"].ToString();
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

                                    //INGRESO A INSIGNIAS
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INSIGNIAS", true, file);

                                    //SELECCION PERSONA
                                    selenium.Click("//span[@id='select2-ctl00_ContenidoPagina_ddlColaborador-container']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@type='search']", Persona);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("PERSONA", true, file);
                                    selenium.Enter("//input[@type='search']");

                                    //MENSAJE
                                    selenium.Click("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']", Mensaje);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MENSAJE", true, file);

                                    selenium.Scroll("//div[@id='1']/div");

                                    //ARRASTRAR INSIGNIA

                                    IWebElement desde = driver2.FindElement(By.XPath("//div[@id='1']/div"));
                                    IWebElement hasta = driver2.FindElement(By.XPath("//div[@id='container-right']"));

                                    Actions action = new Actions(driver2);
                                    action.DragAndDrop(desde, hasta).Perform();

                                    selenium.Screenshot("INSIGNIA ARRASTRADA", true, file);
                                    Thread.Sleep(2000);

                                    //AUTODIDACTA
                                    selenium.Click("//div[@id='divComportamientos']/div[2]/div/label/span");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("AUTODIDACTA", true, file);

                                    //CONFIRMAR AUTODIDACTA
                                    selenium.Scroll("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    selenium.Click("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(4000);

                                    //CONFIRMAR
                                    selenium.Scroll("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Click("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("SATISFACTORIO", true, file);

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
        public void BP_CaracteresEspecialesObservacionesAsignacionInsignias()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_CaracteresEspecialesObservacionesAsignacionInsignias")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Datos Prueba    
                                rows["Persona"].ToString().Length != 0 && rows["Persona"].ToString() != null &&
                                rows["Mensaje"].ToString().Length != 0 && rows["Mensaje"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Persona = rows["Persona"].ToString();
                                string Mensaje = rows["Mensaje"].ToString();
                                string url2 = rows["url2"].ToString();
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
                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A INSIGNIAS
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INSIGNIAS", true, file);

                                    //SELECCION PERSONA
                                    selenium.Click("//span[@id='select2-ctl00_ContenidoPagina_ddlColaborador-container']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@type='search']", Persona);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("PERSONA", true, file);
                                    selenium.Enter("//input[@type='search']");

                                    //MENSAJE
                                    selenium.Click("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']", Mensaje);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MENSAJE CARACTERES ESPECIALES", true, file);

                                    selenium.Scroll("//div[@id='1']/div");

                                    //ARRASTRAR INSIGNIA

                                    IWebElement desde = driver2.FindElement(By.XPath("//div[@id='1']/div"));
                                    IWebElement hasta = driver2.FindElement(By.XPath("//div[@id='container-right']"));

                                    Actions action = new Actions(driver2);
                                    action.DragAndDrop(desde, hasta).Perform();

                                    selenium.Screenshot("INSIGNIA ARRASTRADA", true, file);
                                    Thread.Sleep(2000);

                                    //AUTODIDACTA
                                    selenium.Click("//div[@id='divComportamientos']/div/div/label/span");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("AUTODIDACTA", true, file);

                                    //CONFIRMAR AUTODIDACTA
                                    selenium.Scroll("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(4000);
                                    selenium.Click("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(4000);

                                    //CONFIRMAR
                                    selenium.Scroll("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Click("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("CARACTERES ESPECIALES", true, file);

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
        public void BP_VisualizaciónComportamientosInsignias()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_VisualizaciónComportamientosInsignias")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Datos Prueba    
                                rows["Persona"].ToString().Length != 0 && rows["Persona"].ToString() != null &&
                                rows["Mensaje"].ToString().Length != 0 && rows["Mensaje"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Persona = rows["Persona"].ToString();
                                string Mensaje = rows["Mensaje"].ToString();
                                string url2 = rows["url2"].ToString();
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

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A INSIGNIAS
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INSIGNIAS", true, file);

                                    //SELECCION PERSONA
                                    selenium.Click("//span[@id='select2-ctl00_ContenidoPagina_ddlColaborador-container']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@type='search']", Persona);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("PERSONA", true, file);
                                    selenium.Enter("//input[@type='search']");

                                    //MENSAJE
                                    selenium.Click("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']", Mensaje);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MENSAJE", true, file);

                                    selenium.Scroll("//div[@id='1']/div");

                                    //ARRASTRAR INSIGNIA

                                    IWebElement desde = driver2.FindElement(By.XPath("//div[@id='1']/div"));
                                    IWebElement hasta = driver2.FindElement(By.XPath("//div[@id='container-right']"));

                                    Actions action = new Actions(driver2);
                                    action.DragAndDrop(desde, hasta).Perform();

                                    selenium.Screenshot("INSIGNIA ARRASTRADA", true, file);
                                    Thread.Sleep(2000);

                                    //AUTODIDACTA
                                    selenium.Click("//div[@id='divComportamientos']/div/div/label/span");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("COMPORTAMIENTOS", true, file);

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
        public void BP_CentroVacacionesRolLiderVideoImágenesAdjuntos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_CentroVacacionesRolLiderVideoImágenesAdjuntos")
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
                                string JefeUser1 = rows["JefeUser1"].ToString();
                                string JefePass1 = rows["JefePass1"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFin = rows["FechaFin"].ToString();
                                string FormaPago = rows["FormaPago"].ToString();
                                string FamiliarID = rows["FamiliarID"].ToString();
                                string FamiliarNombre = rows["FamiliarNombre"].ToString();
                                string FamiliarApellido = rows["FamiliarApellido"].ToString();
                                string FamiliarFecha = rows["FamiliarFecha"].ToString();
                                string Prestamo = rows["Prestamo"].ToString();
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

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string username = Environment.UserName;

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    if (database == "SQL")
                                    {

                                        File.Delete(@"C:\Users\" + username + @"\Downloads\Archivo01022022_011334.PNG");
                                        File.Delete(@"C:\Users\" + username + @"\Downloads\ArchivoDoc1.DOCX");
                                        File.Delete(@"C:\Users\" + username + @"\Downloads\ArchivoArchivoPrueba.PDF");

                                    }
                                    else
                                    {

                                        File.Delete(@"C:\Users\" + username + @"\Downloads\Archivoautodidacta (1).JPG");
                                        File.Delete(@"C:\Users\" + username + @"\Downloads\ArchivoPLANTILLA CONEXIÓN VPN SUPERSOCIEDADES AJUSTADO (1) (1).DOCX");
                                        File.Delete(@"C:\Users\" + username + @"\Downloads\ArchivoArchivoPrueba.PDF");

                                    }
                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               
                                    //PARAMETRIZACION 

                                    if (database == "SQL")
                                    {
                                        string familiar = $"delete from BP_DSCEV where ACT_USUA='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(familiar, database, user);

                                        string centro = $"delete from bp_cenva where COD_EMPR= 9 AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(centro, database, user);

                                        string empleado = $"delete from nm_soltr where COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(empleado, database, user);

                                        string jefe1 = $"delete from nm_soltr where COD_RESP = '{JefeUser1}' AND TIP_APLI='Q'";
                                        db.UpdateDeleteInsert(jefe1, database, user);

                                        string jefe2 = $"delete from nm_soltr where COD_RESP = '11' AND TIP_APLI='Q'";
                                        db.UpdateDeleteInsert(jefe2, database, user);
                                    }

                                    if (database == "ORA")
                                    {
                                        string familiar = $"delete from BP_DSCEV where ACT_USUA='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(familiar, database, user);

                                        string centro = $"delete from bp_cenva where COD_EMPR= 421 AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(centro, database, user);

                                        string empleado = $"delete from nm_soltr where COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(empleado, database, user);

                                        string jefe1 = $"delete from nm_soltr where COD_RESP = '{JefeUser1}' AND TIP_APLI='Q'";
                                        db.UpdateDeleteInsert(jefe1, database, user);

                                        string jefe2 = $"delete from nm_soltr where COD_RESP = '45504088' AND TIP_APLI='Q'";
                                        db.UpdateDeleteInsert(jefe2, database, user);
                                    }
                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A CENTRO VACACIONAL
                                    selenium.Scroll("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);

                                    if (database == "ORA")
                                    {
                                        selenium.Scroll("//*[@id='MenuContex']/div[2]/div[1]/ul[1]/li[22]/ul[1]/li[1]/a[1]");
                                        selenium.Click("//*[@id='MenuContex']/div[2]/div[1]/ul[1]/li[22]/ul[1]/li[1]/a[1]");
                                        Thread.Sleep(2000);
                                    }
                                    else
                                    {
                                        selenium.Scroll("//a[contains(@href, 'frmBpCenvacaL.aspx')]");
                                        selenium.Click("//a[contains(@href, 'frmBpCenvacaL.aspx')]");
                                        Thread.Sleep(2000);
                                    }
                                    selenium.Screenshot("MIS CENTROS VACACIONALES", true, file);
                                    //SELECCION CENTRO
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//a/div/div/img");
                                        Thread.Sleep(10000);
                                        selenium.Screenshot("Solicitudes Centros Vacacionales", true, file);
                                    }
                                    else
                                    {
                                        selenium.Click("//a/div/div/img");
                                        Thread.Sleep(10000);
                                        selenium.Screenshot("Solicitudes Centros Vacacionales", true, file);
                                    }

                                    //FECHAS
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFechIni_txtFecha']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFechIni_txtFecha']", FechaInicial);
                                    Thread.Sleep(10000);
                                    selenium.Screenshot("Fecha Inicial", true, file);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFechFin_txtFecha']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFechFin_txtFecha']", FechaFin);
                                    selenium.Screenshot("Fecha Final", true, file);
                                    SendKeys.SendWait("{TAB}");

                                    Thread.Sleep(10000);
                                    //CHECK HACE PARTE GRUPO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_rblUstHues_0']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblUstHues_0']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Parte Acompañantes", true, file);
                                    //FAMILIAR
                                    selenium.Scroll("//div[@id='ctl00_ContenidoPagina_ddlFamil_sl']/div");
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_ddlFamil_sl']/div");
                                    Thread.Sleep(3000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_ddlFamil_0']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Seleccionar Familiar", true, file);
                                    selenium.Click("//input[@value='Aceptar']");
                                    Thread.Sleep(3000);
                                    //FORMA PAGO
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_rdbforpago']");
                                    Thread.Sleep(3000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_rdbforpago']", FormaPago);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Forma Pago Préstamo", true, file);
                                    //TOTAL A PAGAR
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_btnValCan']/span");
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_btnValCan']/span");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Total Pagar", true, file);
                                    //OBSERVACIONES
                                    selenium.Scroll("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValObser_txtTexto']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValObser_txtTexto']", "PRUEBAS CENTROS VACAIONALES");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Observaciones", true, file);
                                   
                                    //ADICIONAR ACOMPAÑANTES
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_OtroAcomp']");
                                    Thread.Sleep(3000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_OtroAcomp']");
                                    Thread.Sleep(3000);
                                    //ID ACOMPAÑANTE
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtidAcomp']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtidAcomp']", FamiliarID);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("ID Acompañante", true, file);
                                    //NOMBRE ACOMPAÑANTE
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtNombre']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNombre']", FamiliarNombre);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nombre Acompañante", true, file);
                                    //APELLIDO ACOMPAÑANTE
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtApellido']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtApellido']", FamiliarApellido);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Apellido Acompañante", true, file);
                                    //FECHA ACOMPAÑANTE
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_KCtrlFecAcomp_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecAcomp_txtFecha']", FamiliarFecha);
                                    Thread.Sleep(3000);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Screenshot("Fecha Nacimiento Acompañante", true, file);
                                    //GUARDAR ACOMPAÑANTE
                                    selenium.Scroll("//a[contains(text(),'Guardar acompañante')]");
                                    selenium.Click("//a[contains(text(),'Guardar acompañante')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Acompañante Agregado", true, file);
                                    //GUARDAR
                                    selenium.Scroll("//a[contains(text(),'Aceptar')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Acompañante Agregado", true, file);
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[contains(text(),'Aceptar')]");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Terminos y Condiciones", true, file);
                                    selenium.Click("//input[@id='ctl00_ContentPopapModel_btnSi']");
                                    Thread.Sleep(5000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContentPopapModel_ddlNomPres']", Prestamo);
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Prestamo", true, file);
                                    selenium.Click("(//a[contains(text(),'Aceptar')])[2]");
                                    Thread.Sleep(10000);
                                    selenium.Screenshot("Registro exitoso", true, file);
                                    selenium.Close();
                                    //--------------------------------------ROL LIDER------------------------------------------------
                                    //LOGIN
                                    selenium.LoginApps(app, JefeUser1, JefePass1, url, file);
                                    Thread.Sleep(2000);

                                    //SOLICITUDES CENTROS VACACIONALES
                                    selenium.Click("//button[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Líder", true, file);
                                    selenium.Scroll("//a[contains(.,'CENTROS VACACIONALES')]");
                                    selenium.Click("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Solicitudes Centros Vacacionales')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Solicitudes Centros Vacacionales por aprobar", true, file);
                                    Thread.Sleep(3000);
                                    //BUSCAR
                                    selenium.SendKeys("//*[@id='tableCentroVaca_filter']/label/input", EmpleadoUser);
                                    Thread.Sleep(6000);
                                    //DETALLE   
                                    selenium.Click("//table[@id='tableCentroVaca']/tbody/tr/td[9]/a/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Solicitud Centro Vacacional a aprobar", true, file);
                                    
                                    //SELECCIONAR VIDEO
                                    selenium.Click("//a[contains(@href, 'https://youtu.be/LcHhMXn5m-ohttps://youtu.be/LcHhMXn5m-o')]");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("VIDEO", true, file);
                                    Thread.Sleep(6000);
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().Refresh();
                                    Thread.Sleep(2000);

                                    //IMAGENES
                                    selenium.Scroll("//img[@id='ctl00_ContenidoPagina_IMGCEN_2']");
                                    selenium.Click("//img[@id='ctl00_ContenidoPagina_IMGCEN_2']");
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_IMGCEN2']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("IMAGEN 1", true, file);

                                    for (int i = 0; i < 2; i++)
                                    {
                                        SendKeys.SendWait("{RIGHT}");
                                        selenium.Screenshot("IMAGENES", true, file);
                                        Thread.Sleep(3000);

                                    }
                                    driver2.Navigate().Refresh();
                                    Thread.Sleep(6000);

                                    //DOCUMENTOS
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocuCenva_ctl02_LinkButton1']/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("DOCUMENTOS", true, file);
                                    Thread.Sleep(3000);

                                    //PDF
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocuCenva_ctl02_LinkButton1']/i");
                                    Thread.Sleep(6000);
                                    //VENTANA DESCARGA PDF
                                    String mainWin = selenium.MainWindow();
                                    String modalWin = selenium.PopupWindow();
                                    Thread.Sleep(3000);
                                    selenium.ChangeWindow(modalWin);
                                    Thread.Sleep(5000);
                                    Screenshot("DESCARGA PDF", true, file);
                                    selenium.Close();
                                    selenium.ChangeWindow(mainWin);


                                    //WORD
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocuCenva_ctl02_LinkButton1']/i");
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocuCenva_ctl03_LinkButton1']/i");
                                    Thread.Sleep(6000);
                                    //VENTANA DESCARGA WORD
                                    String mainWin1 = selenium.MainWindow();
                                    String modalWin1 = selenium.PopupWindow();
                                    Thread.Sleep(3000);
                                    selenium.ChangeWindow(modalWin1);
                                    Thread.Sleep(5000);
                                    Screenshot("DESCARGA WORD", true, file);
                                    selenium.Close();
                                    selenium.ChangeWindow(mainWin1);

                                    //IMAGEN
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocuCenva_ctl02_LinkButton1']/i");
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocuCenva_ctl04_LinkButton1']/i");
                                    Thread.Sleep(6000);
                                    //VENTANA DESCARGA imagen
                                    String mainWin2 = selenium.MainWindow();
                                    String modalWin2 = selenium.PopupWindow();
                                    Thread.Sleep(5000);
                                    selenium.ChangeWindow(modalWin2);
                                    Thread.Sleep(4000);
                                    Screenshot("DESCARGA IMAGEN", true, file);
                                    selenium.Close();
                                    selenium.ChangeWindow(mainWin2);
                                    if (database == "SQL")
                                    {
                                        //ABRIR IMAGEN DESCARGADO
                                        string imagenPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\Archivo01022022_011334.PNG");
                                        Process.Start(imagenPath);
                                        Thread.Sleep(10000);
                                        Screenshot("IMAGEN ABIERTA", true, file);

                                        //ABRIR PDF DESCARGADO
                                        string pdfPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\ArchivoArchivoPrueba.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(10000);
                                        Screenshot("PDF ABIERTO", true, file);

                                        //ABRIR WORD DESCARGADO
                                        string wordPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\ArchivoDoc1.DOCX");
                                        Process.Start(wordPath);
                                        Thread.Sleep(10000);
                                        Screenshot("WORD ABIERTO", true, file);

                                    }
                                    else
                                    {
                                        //ABRIR IMAGEN DESCARGADO
                                        string imagenPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\Archivoautodidacta.JPG");
                                        Process.Start(imagenPath);
                                        Thread.Sleep(10000);
                                        Screenshot("IMAGEN ABIERTA", true, file);

                                        //ABRIR PDF DESCARGADO
                                        string pdfPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\ArchivoArchivoPrueba.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(10000);
                                        Screenshot("PDF ABIERTO", true, file);

                                        //ABRIR WORD DESCARGADO
                                        string wordPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\ArchivoPLANTILLA CONEXIÓN VPN SUPERSOCIEDADES AJUSTADO (1) (1).DOCX");
                                        Process.Start(wordPath);
                                        Thread.Sleep(10000);
                                        Screenshot("WORD ABIERTO", true, file);

                                    }

                                    Thread.Sleep(5000);
                                    selenium.Close();
                                    KillProcesos("Acrobat.exe");
                                    Thread.Sleep(5000);
                                    KillProcesos("WINWORD.EXE");
                                    Thread.Sleep(5000);
                                    KillProcesos("Microsoft.Photos.exe");
                                    Thread.Sleep(5000);
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
        public void BP_ValidarModuloDeReconocimientosEnN()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_ValidarModuloDeReconocimientosEnN")
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

                                        string actualizar = $"update gn_modul set INS_MODU = 'N' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

                                    }
                                    else
                                    {
                                        string actualizar = $"update gn_modul set INS_MODU = 'N' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 507195 AND COD_EMPR = 9";
                                        db.UpdateDeleteInsert(actualizar1, database, user);
                                    }

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MODULO RECONOCIMIENTO NO VISIBLE", true, file);
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
        public void BP_ValidarModuloDeReconocimientosEnS()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_ValidarModuloDeReconocimientosEnS")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba    
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string url2 = rows["url2"].ToString();
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

                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

                                    }
                                    else
                                    {
                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 507195 AND COD_EMPR = 9";
                                        db.UpdateDeleteInsert(actualizar1, database, user);
                                    }

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MODULO RECONOCIMIENTO AVTIVADO", true, file);
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INSIGNIAS", true, file);
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
        public void BP_ValidacionDatosBasicosInsignias()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_ValidacionDatosBasicosInsignias")
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
                                        string actualizar = $"update nm_contr set cod_Carg = 'ABC123' where COD_EMPR = 421 AND COD_EMPL = 19301797 and nro_cont = 1";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                        string actualizar1 = $"update bi_emple set box_mail= 'deisyl@digitalware.com.co', dir_resi = 'AK 70O BIS N SUR68Q UR', tel_movi = 3143142468 where COD_EMPR = 421 AND COD_EMPL = 19301797";
                                        db.UpdateDeleteInsert(actualizar1, database, user);
                                        string actualizar2 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar2, database, user);
                                        string actualizar3 = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar3, database, user);

                                    }
                                    else
                                    {
                                        string actualizar = $"update nm_contr set cod_Carg = '1' where COD_EMPR = 9 AND COD_EMPL = 507195 and nro_cont = 1";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                        string actualizar1 = $"update bi_emple set box_mail= 'deisyl@digitalware.com.co', dir_resi = 'AK 3 20 78', tel_movi = 3143142443 where COD_EMPR = 9 AND COD_EMPL = 507195";
                                        db.UpdateDeleteInsert(actualizar1, database, user);
                                        string actualizar2 = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar2, database, user);
                                        string actualizar3 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 507195 AND COD_EMPR = 9";
                                        db.UpdateDeleteInsert(actualizar3, database, user);
                                    }

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login", true, file);

                                    //MIS INSIGNIAS
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Mis Insignias')]");
                                    selenium.Click("//a[contains(.,'Mis Insignias')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MIs Insignias", true, file);

                                    //Datos Basicos
                                    selenium.Screenshot("Datos Basicos", true, file);
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
        public void BP_ValidarIndicadorActividadkbpinsig()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_ValidarIndicadorActividadkbpinsig")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba    
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string url2 = rows["url2"].ToString();
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

                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

                                    }
                                    else
                                    {
                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 507195 AND COD_EMPR = 9";
                                        db.UpdateDeleteInsert(actualizar1, database, user);
                                    }

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login", true, file);
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Nuestras Insignias", true, file);
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
        public void BP_VisualizacionIconosInsignias()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_VisualizacionIconosInsignias")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba    
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string url2 = rows["url2"].ToString();
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

                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

                                    }
                                    else
                                    {
                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 507195 AND COD_EMPR = 9";
                                        db.UpdateDeleteInsert(actualizar1, database, user);
                                    }

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login", true, file);
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    selenium.ScrollTo("0", "400");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Visualizacion Iconos", true, file);
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
        public void BP_VisualizacionImágenesInsignias()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_VisualizacionImágenesInsignias")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba    
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string url2 = rows["url2"].ToString();
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

                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 19301797 AND COD_EMPR = 421";
                                        db.UpdateDeleteInsert(actualizar1, database, user);

                                    }
                                    else
                                    {
                                        string actualizar = $"update gn_modul set INS_MODU = 'S' where ini_modu = 'RE'";
                                        db.UpdateDeleteInsert(actualizar, database, user);

                                        string actualizar1 = $"update NM_CONTR set IND_aCTI = 'A' where COD_EMPL = 507195 AND COD_EMPR = 9";
                                        db.UpdateDeleteInsert(actualizar1, database, user);
                                    }


                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login", true, file);
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Visualizacion Imagenes Insignias", true, file);
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
        public void BP_InserciónEnkbpasins()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_InserciónEnkbpasins")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Datos Prueba    
                                rows["Persona"].ToString().Length != 0 && rows["Persona"].ToString() != null &&
                                rows["Mensaje"].ToString().Length != 0 && rows["Mensaje"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Persona = rows["Persona"].ToString();
                                string Mensaje = rows["Mensaje"].ToString();
                                string url2 = rows["url2"].ToString();
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
                                        string borrar = $"delete from BP_ASINS where cod_empr = 421 and cod_empl = 55454 And cod_regi = 'B' and tip_apli = 'AC' AND COD_ROLI = 'G' AND COD_ASIG = 19301797";
                                        db.UpdateDeleteInsert(borrar, database, user);
                                    }
                                    else
                                    {
                                        string borrar = $"delete from BP_ASINS where cod_empr = 9 and cod_empl = 13 and cod_insi = '1' And cod_regi = '1'and tip_apli = 'AC' and COD_ASIG = 507195";
                                        db.UpdateDeleteInsert(borrar, database, user);
                                    }


                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A INSIGNIAS
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INSIGNIAS", true, file);

                                    //SELECCION PERSONA
                                    selenium.Click("//span[@id='select2-ctl00_ContenidoPagina_ddlColaborador-container']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@type='search']", Persona);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("PERSONA", true, file);
                                    selenium.Enter("//input[@type='search']");

                                    //MENSAJE
                                    selenium.Click("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']", Mensaje);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MENSAJE", true, file);

                                    selenium.Scroll("//div[@id='1']/div");

                                    //ARRASTRAR INSIGNIA

                                    IWebElement desde = driver2.FindElement(By.XPath("//div[@id='1']/div"));
                                    IWebElement hasta = driver2.FindElement(By.XPath("//div[@id='container-right']"));

                                    Actions action = new Actions(driver2);
                                    action.DragAndDrop(desde, hasta).Perform();

                                    selenium.Screenshot("INSIGNIA ARRASTRADA", true, file);
                                    Thread.Sleep(2000);

                                    //AUTODIDACTA
                                    selenium.Click("//div[@id='divComportamientos']/div/div/label/span");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("COMPORTAMIENTOS", true, file);

                                    //CONFIRMAR AUTODIDACTA
                                    selenium.Scroll("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    selenium.Click("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(4000);

                                    //CONFIRMAR
                                    selenium.Scroll("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Click("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("SATISFACTORIO", true, file);


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
        public void BP_ValidarCantidadMensualkbpinsigDescontando()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_ValidarCantidadMensualkbpinsigDescontando")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                //Datos Prueba    
                                rows["Persona"].ToString().Length != 0 && rows["Persona"].ToString() != null &&
                                rows["Mensaje"].ToString().Length != 0 && rows["Mensaje"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Persona = rows["Persona"].ToString();
                                string Mensaje = rows["Mensaje"].ToString();
                                string url2 = rows["url2"].ToString();
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
                                        string borrar = $"delete from BP_ASINS where cod_empr = 421 and cod_empl = 19301797 And cod_regi = 'B' and cod_insi = '1' AND COD_ROLI = 'A' AND COD_ASIG = 39801386";
                                        db.UpdateDeleteInsert(borrar, database, user);
                                    }
                                    else
                                    {
                                        string borrar = $"delete from BP_ASINS where cod_empr = 9 and cod_empl = 13 and cod_insi = '1' And cod_regi = '1'and tip_apli = 'AC' and COD_ASIG = 507195";
                                        db.UpdateDeleteInsert(borrar, database, user);
                                    }


                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A INSIGNIAS
                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INSIGNIAS", true, file);

                                    //SELECCION PERSONA
                                    selenium.Click("//span[@id='select2-ctl00_ContenidoPagina_ddlColaborador-container']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@type='search']", Persona);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("PERSONA", true, file);
                                    selenium.Enter("//input[@type='search']");

                                    //MENSAJE
                                    selenium.Click("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KctMensaje_txtTexto']", Mensaje);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MENSAJE", true, file);

                                    selenium.Scroll("//div[@id='1']/div");

                                    //ARRASTRAR INSIGNIA

                                    IWebElement desde = driver2.FindElement(By.XPath("//div[@id='1']/div"));
                                    IWebElement hasta = driver2.FindElement(By.XPath("//div[@id='container-right']"));

                                    Actions action = new Actions(driver2);
                                    action.DragAndDrop(desde, hasta).Perform();

                                    selenium.Screenshot("INSIGNIA ARRASTRADA", true, file);
                                    Thread.Sleep(2000);

                                    //AUTODIDACTA
                                    selenium.Click("//div[@id='divComportamientos']/div/div/label/span");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("COMPORTAMIENTOS", true, file);

                                    //CONFIRMAR AUTODIDACTA

                                    selenium.Click("//div[@id='modalComportamientos']/div/div/div[2]/button[2]");
                                    Thread.Sleep(4000);

                                    //CONFIRMAR
                                    selenium.Scroll("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Click("/html/body/form/div[3]/div[3]/div[2]/div/div[2]/div/div/div[1]/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/button");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("SATISFACTORIO", true, file);
                                    Thread.Sleep(4000);
                                    SendKeys.SendWait("{F5}");
                                    Thread.Sleep(8000);
                                    selenium.Scroll("//div[@id='1']/div");
                                    selenium.Screenshot("Insignia Descontada", true, file);
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
        public void BP_VisualizacionInsignias()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_VisualizacionInsignias")
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

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login", true, file);

                                    //MIS INSIGNIAS
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Mis Insignias')]");
                                    selenium.Click("//a[contains(.,'Mis Insignias')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("MIs Insignias", true, file);

                                    //MIs insignias
                                    selenium.Screenshot("Mis Insignias", true, file);
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
        public void BP_ValidacionRankingCompañía()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_ValidacionRankingCompañía")
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
                                rows["Compañia"].ToString().Length != 0 && rows["Compañia"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                //Datos Prueba    


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Compañia = rows["Compañia"].ToString();
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


                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login", true, file);

                                    //MIS INSIGNIAS
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Ranking Insignias')]");
                                    selenium.Click("//a[contains(.,'Ranking Insignias')]");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Ranking Insignias", true, file);

                                    //Compañia
                                    selenium.SelectElementByName("//select[@id='ddlTipoRanking']", Compañia);
                                    Thread.Sleep(10000);
                                    selenium.ScrollTo("0", "200");
                                    selenium.Screenshot("Ranking Compañia", true, file);
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
        public void BP_ValidacionRankingColaboradoresAutodidacta()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_ValidacionRankingColaboradoresAutodidacta")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba    
                                rows["Colaboradores"].ToString().Length != 0 && rows["Colaboradores"].ToString() != null &&
                                rows["Insignia"].ToString().Length != 0 && rows["Insignia"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Colaboradores = rows["Colaboradores"].ToString();
                                string Insignia = rows["Insignia"].ToString();
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

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login", true, file);

                                    //MIS INSIGNIAS
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Ranking Insignias')]");
                                    selenium.Click("//a[contains(.,'Ranking Insignias')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Ranking Insignias", true, file);

                                    //Compañia
                                    selenium.SelectElementByName("//select[@id='ddlTipoRanking']", Colaboradores);
                                    selenium.Screenshot("Ranking Colaboradores", true, file);

                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlInsig']", Insignia);
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Ranking Colaboradores Autodidacta", true, file);
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
        public void BP_ValidacionRankingColaboradoresEmprendedor()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_ValidacionRankingColaboradoresEmprendedor")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba    
                                rows["Colaboradores"].ToString().Length != 0 && rows["Colaboradores"].ToString() != null &&
                                rows["Insignia"].ToString().Length != 0 && rows["Insignia"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Colaboradores = rows["Colaboradores"].ToString();
                                string Insignia = rows["Insignia"].ToString();
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

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login", true, file);

                                    //MIS INSIGNIAS
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Ranking Insignias')]");
                                    selenium.Click("//a[contains(.,'Ranking Insignias')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Ranking Insignias", true, file);

                                    //Compañia
                                    selenium.SelectElementByName("//select[@id='ddlTipoRanking']", Colaboradores);
                                    selenium.Screenshot("Ranking Colaboradores", true, file);

                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlInsig']", Insignia);
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Ranking Colaboradores Emprendedor", true, file);
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
        public void BP_ValidacionRankingAreas()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_ValidacionRankingAreas")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Datos Prueba    
                                rows["Area"].ToString().Length != 0 && rows["Area"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Area = rows["Area"].ToString();

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

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Login", true, file);

                                    //MIS INSIGNIAS
                                    selenium.Click("//a[contains(.,'MI PUESTO DE TRABAJO')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Ranking Insignias')]");
                                    selenium.Click("//a[contains(.,'Ranking Insignias')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Ranking Insignias", true, file);

                                    //Compañia
                                    selenium.SelectElementByName("//select[@id='ddlTipoRanking']", Area);
                                    Thread.Sleep(20000);
                                    selenium.Screenshot("Ranking Area Emprendedor", true, file);
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
        public void BP_CentroVacacionesRolColaboradorVideoImágenesAdjuntos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_CentroVacacionesRolColaboradorVideoImágenesAdjuntos")
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
                               rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string url2 = rows["url2"].ToString();
                                string Ruta1 = rows["Ruta1"].ToString();
                                string Ruta2 = rows["Ruta2"].ToString();
                                string Ruta3 = rows["Ruta3"].ToString();

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

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string username = Environment.UserName;
                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    if (database == "SQL")
                                    {

                                        File.Delete(@"C:\Users\" + username + @"\Downloads\Archivo01022022_011334.PNG");
                                        File.Delete(@"C:\Users\" + username + @"\Downloads\ArchivoDoc1.DOCX");
                                        File.Delete(@"C:\Users\" + username + @"\Downloads\ArchivoArchivoPrueba.PDF");

                                    }
                                    else
                                    {

                                        File.Delete(@"C:\Users\" + username + @"\Downloads\Archivoautodidacta.JPG");
                                        File.Delete(@"C:\Users\" + username + @"\Downloads\ArchivoPLANTILLA CONEXIÓN VPN SUPERSOCIEDADES AJUSTADO (1) (1).DOCX");
                                        File.Delete(@"C:\Users\" + username + @"\Downloads\ArchivoArchivoPrueba.PDF");

                                    }
                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A CENTRO VACACIONAL
                                    selenium.Scroll("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);
                                    if (database == "ORA")
                                    {
                                        selenium.Scroll("//*[@id='MenuContex']/div[2]/div[1]/ul[1]/li[22]/ul[1]/li[1]/a[1]");
                                        selenium.Click("//*[@id='MenuContex']/div[2]/div[1]/ul[1]/li[22]/ul[1]/li[1]/a[1]");
                                        Thread.Sleep(2000);
                                    }
                                    else
                                    {
                                        selenium.Scroll("//a[contains(@href, 'frmBpCenvacaL.aspx')]");
                                        selenium.Click("//a[contains(@href, 'frmBpCenvacaL.aspx')]");
                                        Thread.Sleep(2000);
                                    }
                                    selenium.Screenshot("MIS CENTROS VACACIONALES", true, file);
                                    //SELECCION CENTRO
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//a/div/div/img");
                                        Thread.Sleep(3000);
                                    }
                                    else
                                    {
                                        selenium.Click("//a/div/div/img");
                                        Thread.Sleep(3000);
                                    }
                                    //SELECCIONAR VIDEO
                                    selenium.Click("//a[contains(@href, 'https://youtu.be/LcHhMXn5m-ohttps://youtu.be/LcHhMXn5m-o')]");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("VIDEO", true, file);
                                    Thread.Sleep(6000);

                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().Refresh();
                                    Thread.Sleep(2000);


                                    //IMAGENES
                                    selenium.Scroll("//img[@id='ctl00_ContenidoPagina_IMGCEN_2']");
                                    Thread.Sleep(3000);
                                    selenium.Click("//img[@id='ctl00_ContenidoPagina_IMGCEN_2']");
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_IMGCEN2']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("IMAGEN 1", true, file);

                                    for (int i = 0; i < 2; i++)
                                    {
                                        SendKeys.SendWait("{RIGHT}");
                                        selenium.Screenshot("IMAGENES", true, file);
                                        Thread.Sleep(3000);

                                    }
                                    driver2.Navigate().Refresh();
                                    Thread.Sleep(6000);

                                    //DOCUMENTOS
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocuCenva_ctl02_LinkButton1']/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("DOCUMENTOS", true, file);
                                    Thread.Sleep(3000);

                                    //PDF
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocuCenva_ctl02_LinkButton1']/i");
                                    Thread.Sleep(10000);
                                    //VENTANA DESCARGA PDF
                                    String mainWin = selenium.MainWindow();
                                    String modalWin = selenium.PopupWindow();
                                    Thread.Sleep(3000);
                                    selenium.ChangeWindow(modalWin);
                                    Thread.Sleep(5000);

                                    Screenshot("DESCARGA PDF", true, file);
                                    selenium.Close();
                                    selenium.ChangeWindow(mainWin);


                                    //WORD
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocuCenva_ctl02_LinkButton1']/i");
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocuCenva_ctl03_LinkButton1']/i");
                                    Thread.Sleep(10000);
                                    //VENTANA DESCARGA WORD
                                    String mainWin1 = selenium.MainWindow();
                                    String modalWin1 = selenium.PopupWindow();
                                    Thread.Sleep(3000);
                                    selenium.ChangeWindow(modalWin1);
                                    Thread.Sleep(5000);

                                    Screenshot("DESCARGA WORD", true, file);
                                    selenium.Close();
                                    selenium.ChangeWindow(mainWin1);

                                    //IMAGEN
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocuCenva_ctl02_LinkButton1']/i");
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgGnGestDocuCenva_ctl04_LinkButton1']/i");
                                    Thread.Sleep(10000);
                                    //VENTANA DESCARGA imagen
                                    String mainWin2 = selenium.MainWindow();
                                    String modalWin2 = selenium.PopupWindow();
                                    Thread.Sleep(5000);
                                    selenium.ChangeWindow(modalWin2);
                                    Thread.Sleep(4000);

                                    Screenshot("DESCARGA IMAGEN", true, file);
                                    selenium.Close();
                                    selenium.ChangeWindow(mainWin2);

                                    if (database == "SQL")
                                    {
                                        //ABRIR IMAGEN DESCARGADO
                                        string imagenPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\Archivo01022022_011334.PNG");
                                        Process.Start(imagenPath);
                                        Thread.Sleep(10000);
                                        Screenshot("IMAGEN ABIERTA", true, file);

                                        //ABRIR PDF DESCARGADO
                                        string pdfPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\ArchivoArchivoPrueba.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(10000);
                                        Screenshot("PDF ABIERTO", true, file);

                                        //ABRIR WORD DESCARGADO
                                        string wordPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\ArchivoDoc1.DOCX");
                                        Process.Start(wordPath);
                                        Thread.Sleep(10000);
                                        Screenshot("WORD ABIERTO", true, file);

                                    }
                                    else
                                    {
                                        //ABRIR IMAGEN DESCARGADO
                                        string imagenPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\Archivoautodidacta.JPG");
                                        Process.Start(imagenPath);
                                        Thread.Sleep(10000);
                                        Screenshot("IMAGEN ABIERTA", true, file);

                                        //ABRIR PDF DESCARGADO
                                        string pdfPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\ArchivoArchivoPrueba.PDF");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(10000);
                                        Screenshot("PDF ABIERTO", true, file);

                                        //ABRIR WORD DESCARGADO
                                        string wordPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\ArchivoPLANTILLA CONEXIÓN VPN SUPERSOCIEDADES AJUSTADO (1) (1).DOCX");
                                        Process.Start(wordPath);
                                        Thread.Sleep(10000);
                                        Screenshot("WORD ABIERTO", true, file);

                                    }
                                    selenium.Close();
                                    Thread.Sleep(4000);
                                    LimpiarProcesos();
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
        public void BP_CentrosVacacionalesFormaPagoPrestamosAutorizaciónJefeEspecificoJefedelJefe()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BP.BP_CentrosVacacionalesFormaPagoPrestamosAutorizaciónJefeEspecificoJefedelJefe")
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
                                rows["JefeUser1"].ToString().Length != 0 && rows["JefeUser1"].ToString() != null &&
                                rows["JefePass1"].ToString().Length != 0 && rows["JefePass1"].ToString() != null &&
                                rows["JefeUser2"].ToString().Length != 0 && rows["JefeUser2"].ToString() != null &&
                                rows["JefePass2"].ToString().Length != 0 && rows["JefePass2"].ToString() != null &&
                                rows["FechaInicial"].ToString().Length != 0 && rows["FechaInicial"].ToString() != null &&
                                rows["FechaFin"].ToString().Length != 0 && rows["FechaFin"].ToString() != null &&
                                rows["FormaPago"].ToString().Length != 0 && rows["FormaPago"].ToString() != null &&
                                rows["FamiliarID"].ToString().Length != 0 && rows["FamiliarID"].ToString() != null &&
                                rows["FamiliarNombre"].ToString().Length != 0 && rows["FamiliarNombre"].ToString() != null &&
                                rows["FamiliarApellido"].ToString().Length != 0 && rows["FamiliarApellido"].ToString() != null &&
                                rows["FamiliarFecha"].ToString().Length != 0 && rows["FamiliarFecha"].ToString() != null &&
                                rows["Prestamo"].ToString().Length != 0 && rows["Prestamo"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string JefeUser1 = rows["JefeUser1"].ToString();
                                string JefePass1 = rows["JefePass1"].ToString();
                                string JefeUser2 = rows["JefeUser2"].ToString();
                                string JefePass2 = rows["JefePass2"].ToString();
                                string FechaInicial = rows["FechaInicial"].ToString();
                                string FechaFin = rows["FechaFin"].ToString();
                                string FormaPago = rows["FormaPago"].ToString();
                                string FamiliarID = rows["FamiliarID"].ToString();
                                string FamiliarNombre = rows["FamiliarNombre"].ToString();
                                string FamiliarApellido = rows["FamiliarApellido"].ToString();
                                string FamiliarFecha = rows["FamiliarFecha"].ToString();
                                string Prestamo = rows["Prestamo"].ToString();
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

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               
                                    //PARAMETRIZACION 
                                    if (database == "SQL")
                                    {
                                        string familiar = $"delete from BP_DSCEV where ACT_USUA='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(familiar, database, user);

                                        string centro = $"delete from bp_cenva where COD_EMPR= 9 AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(centro, database, user);

                                        string empleado = $"delete from nm_soltr where COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(empleado, database, user);

                                        string jefe1 = $"delete from nm_soltr where COD_RESP = '{JefeUser1}' AND TIP_APLI='Q'";
                                        db.UpdateDeleteInsert(jefe1, database, user);

                                        string jefe2 = $"delete from nm_soltr where COD_RESP = '{JefeUser2}' AND TIP_APLI='Q'";
                                        db.UpdateDeleteInsert(jefe2, database, user);
                                    }
                                    
                                    if (database == "ORA")
                                    {
                                        string familiar = $"delete from BP_DSCEV where ACT_USUA='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(familiar, database, user);

                                        string centro = $"delete from bp_cenva where COD_EMPR= 421 AND COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(centro, database, user);

                                        string empleado = $"delete from nm_soltr where COD_RESP ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(empleado, database, user);

                                        string jefe1 = $"delete from nm_soltr where COD_RESP = '{JefeUser1}' AND TIP_APLI='Q'";
                                        db.UpdateDeleteInsert(jefe1, database, user);

                                        string jefe2 = $"delete from nm_soltr where COD_RESP = '{JefeUser2}' AND TIP_APLI='Q'";
                                        db.UpdateDeleteInsert(jefe2, database, user);
                                    }
                                    


                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A CENTRO VACACIONAL
                                    selenium.Scroll("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);
                                    
                                    if (database == "ORA")
                                    {
                                        selenium.Scroll("//*[@id='MenuContex']/div[2]/div[1]/ul[1]/li[22]/ul[1]/li[1]/a[1]");
                                        selenium.Click("//*[@id='MenuContex']/div[2]/div[1]/ul[1]/li[22]/ul[1]/li[1]/a[1]");
                                        Thread.Sleep(2000);
                                    }
                                    else
                                    {
                                        selenium.Scroll("//a[contains(@href, 'frmBpCenvacaL.aspx')]");
                                        selenium.Click("//a[contains(@href, 'frmBpCenvacaL.aspx')]");
                                        Thread.Sleep(2000);
                                    }
                                    selenium.Screenshot("MIS CENTROS VACACIONALES", true, file);
                                    //SELECCION CENTRO
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//a/div/div/img");
                                        Thread.Sleep(10000);
                                        selenium.Screenshot("Solicitudes Centros Vacacionales", true, file);
                                    }
                                    else
                                    {
                                        selenium.Click("//a/div/div/img");
                                        Thread.Sleep(10000);
                                        selenium.Screenshot("Solicitudes Centros Vacacionales", true, file);
                                    }

                                    //FECHAS
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFechIni_txtFecha']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFechIni_txtFecha']", FechaInicial);
                                    Thread.Sleep(10000);
                                    selenium.Screenshot("Fecha Inicial", true, file);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_KCtrlFechFin_txtFecha']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFechFin_txtFecha']", FechaFin);
                                    selenium.Screenshot("Fecha Final", true, file);
                                    SendKeys.SendWait("{TAB}");

                                    Thread.Sleep(10000);
                                    //CHECK HACE PARTE GRUPO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_rblUstHues_0']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblUstHues_0']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Parte Acompañantes", true, file);
                                    //FAMILIAR
                                    selenium.Scroll("//div[@id='ctl00_ContenidoPagina_ddlFamil_sl']/div");
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_ddlFamil_sl']/div");
                                    Thread.Sleep(3000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_ddlFamil_0']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Seleccionar Familiar", true, file);
                                    selenium.Click("//input[@value='Aceptar']");
                                    Thread.Sleep(3000);
                                    //FORMA PAGO
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_rdbforpago']");
                                    Thread.Sleep(3000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_rdbforpago']", FormaPago);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Forma Pago Préstamo", true, file);
                                    //TOTAL A PAGAR
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_btnValCan']/span");
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_btnValCan']/span");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Total Pagar", true, file);
                                    //OBSERVACIONES
                                    selenium.Scroll("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValObser_txtTexto']");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValObser_txtTexto']", "PRUEBAS CENTROS VACAIONALES");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Observaciones", true, file);

                                    //ADICIONAR ACOMPAÑANTES
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_OtroAcomp']");
                                    Thread.Sleep(3000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_OtroAcomp']");
                                    Thread.Sleep(3000);
                                    //ID ACOMPAÑANTE
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtidAcomp']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtidAcomp']", FamiliarID);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("ID Acompañante", true, file);
                                    //NOMBRE ACOMPAÑANTE
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtNombre']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNombre']", FamiliarNombre);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Nombre Acompañante", true, file);
                                    //APELLIDO ACOMPAÑANTE
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtApellido']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtApellido']", FamiliarApellido);
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Apellido Acompañante", true, file);
                                    //FECHA ACOMPAÑANTE
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_KCtrlFecAcomp_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecAcomp_txtFecha']", FamiliarFecha);
                                    Thread.Sleep(3000);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Screenshot("Fecha Nacimiento Acompañante", true, file);
                                    //GUARDAR ACOMPAÑANTE
                                    selenium.Scroll("//a[contains(text(),'Guardar acompañante')]");
                                    selenium.Click("//a[contains(text(),'Guardar acompañante')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Acompañante Agregado", true, file);
                                    //GUARDAR
                                    selenium.Scroll("//a[contains(text(),'Aceptar')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Acompañante Agregado", true, file);
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[contains(text(),'Aceptar')]");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Terminos y Condiciones", true, file);
                                    selenium.Click("//input[@id='ctl00_ContentPopapModel_btnSi']");
                                    Thread.Sleep(5000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContentPopapModel_ddlNomPres']", Prestamo);
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Prestamo", true, file);
                                    selenium.Click("(//a[contains(text(),'Aceptar')])[2]");
                                    Thread.Sleep(10000);
                                    selenium.Screenshot("Registro exitoso", true, file);
                                    selenium.Close();
                                    //--------------------------------------APROBACION ESPECIFICO------------------------------------------------
                                    //LOGIN
                                    selenium.LoginApps(app, JefeUser1, JefePass1, url, file);
                                    Thread.Sleep(2000);

                                    //SOLICITUDES CENTROS VACACIONALES
                                    selenium.Click("//button[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Líder", true, file);
                                    selenium.Scroll("//a[contains(.,'CENTROS VACACIONALES')]");
                                    selenium.Click("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Solicitudes Centros Vacacionales')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Solicitudes Centros Vacacionales por aprobar", true, file);
                                    Thread.Sleep(3000);
                                    //BUSCAR
                                    selenium.SendKeys("//*[@id='tableCentroVaca_filter']/label/input", EmpleadoUser);
                                    Thread.Sleep(6000);
                                    //DETALLE   
                                    selenium.Click("//table[@id='tableCentroVaca']/tbody/tr/td[9]/a/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Solicitud Centro Vacacional a aprobar", true, file);
                                    //APROBAR
                                    selenium.Scroll("//a[contains(text(),'Aprobar')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Aprobar", true, file);
                                    selenium.Click("//a[contains(text(),'Aprobar')]");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Solicitud Aprobada", true, file);
                                    selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    Thread.Sleep(3000);
                                    selenium.Close();

                                    //------------------------------------------APROBADOR JEFE DEJ JEFE-------------------------
                                    //LOGIN
                                    selenium.LoginApps(app, JefeUser2, JefePass2, url, file);
                                    Thread.Sleep(2000);

                                    //SOLICITUDES CENTROS VACACIONALES
                                    selenium.Click("//button[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Líder", true, file);
                                    selenium.Scroll("//a[contains(.,'CENTROS VACACIONALES')]");
                                    selenium.Click("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Solicitudes Centros Vacacionales')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Solicitudes Centros Vacacionales por aprobar", true, file);
                                    Thread.Sleep(3000);
                                    //BUSCAR
                                    selenium.SendKeys("//*[@id='tableCentroVaca_filter']/label/input", EmpleadoUser);
                                    Thread.Sleep(6000);
                                    //DETALLE   
                                    selenium.Click("//table[@id='tableCentroVaca']/tbody/tr/td[9]/a/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Solicitud Centro Vacacional a aprobar", true, file);
                                    //APROBAR
                                    selenium.Scroll("//a[contains(text(),'Aprobar')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Aprobar", true, file);
                                    selenium.Click("//a[contains(text(),'Aprobar')]");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Solicitud Aprobada", true, file);
                                    selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    Thread.Sleep(3000);
                                    selenium.Close();

                                    //------------------------VERIFICAR ESTADO APROBADO SOLICITUD----------------------------------
                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A CENTRO VACACIONAL
                                    selenium.Scroll("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'CENTROS VACACIONALES')]");
                                    Thread.Sleep(2000);

                                    if (database == "ORA")
                                    {
                                        selenium.Scroll("//*[@id='MenuContex']/div[2]/div[1]/ul[1]/li[22]/ul[1]/li[1]/a[1]");
                                        selenium.Click("//*[@id='MenuContex']/div[2]/div[1]/ul[1]/li[22]/ul[1]/li[1]/a[1]");
                                        Thread.Sleep(2000);
                                    }
                                    else
                                    {
                                        selenium.Scroll("//a[contains(@href, 'frmBpCenvacaL.aspx')]");
                                        selenium.Click("//a[contains(@href, 'frmBpCenvacaL.aspx')]");
                                        Thread.Sleep(2000);
                                    }
                                    selenium.Screenshot("Solicitud Aprobada", true, file);
                                    selenium.Close();

                                    fv.ConvertWordToPDF(file, database); 
                                    Thread.Sleep(5000);
                                    string verA = string.Empty;
                                    string verB = string.Empty;
                                    string verC = string.Empty;
                                    string verD = string.Empty;
                                    string verE = string.Empty;
                                    string verF = string.Empty;
                                    string total= string.Empty;
                                    string total1= string.Empty;
                                    //CONSULTA SQL
                                    string consulta = $"SELECT COD_EMPR,COD_EMPL,EST_SOLI FROM BP_CENVA WHERE COD_EMPR= 9 AND COD_EMPL = '4' AND EST_SOLI = 'A'";
                                    DataTable resultado = db.Select(consulta, user, database);
                                    if (resultado.Rows.Count > 0)
                                    {
                                        foreach (DataRow rw in resultado.Rows)
                                        {
                                            verA = rw["COD_EMPR"].ToString();
                                            verB = rw["COD_EMPL"].ToString();
                                            verC = rw["EST_SOLI"].ToString();

                                             total = verA +","+ verB + ","+ verC;
                                        }
                                    }
                                    //CONSULTA SQL
                                    string consulta1 = $"SELECT COD_EMPR,COD_RESP,TIP_APLI FROM NM_SOLTR WHERE COD_EMPR = 9 AND COD_RESP= '4' AND TIP_APLI = 'Q'";
                                    DataTable resultado1 = db.Select(consulta1, user, database);
                                    if (resultado1.Rows.Count > 0)
                                    {
                                        foreach (DataRow rw in resultado1.Rows)
                                        {
                                            verD = rw["COD_EMPR"].ToString();
                                            verE = rw["COD_RESP"].ToString();
                                            verF = rw["TIP_APLI"].ToString();

                                             total1 = verD + "," + verE + "," + verF;
                                        }
                                    }

                                    APIFuncionesVitales.InsertConsulta(file, total, consulta);
                                    APIFuncionesVitales.InsertConsulta(file, total1, consulta1);

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

