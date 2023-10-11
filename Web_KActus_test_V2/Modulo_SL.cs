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
    public class Modulo_SL : FuncionesVitales
    {

        string Modulo = "Modulo_SL";
        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();

        public Modulo_SL()
        {

        }


        [TestMethod]
        public void SL_AprobacionJefeRequisicionDePersonalEmpresa()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_AprobacionJefeRequisicionDePersonalEmpresa")
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
                                //Datos Login Requisición Personal 
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                //Datos Requisición Personal 
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["GrupoRequisiciones"].ToString().Length != 0 && rows["GrupoRequisiciones"].ToString() != null &&
                                rows["FormaCobertura"].ToString().Length != 0 && rows["FormaCobertura"].ToString() != null &&
                                rows["FiltroSeleccion"].ToString().Length != 0 && rows["FiltroSeleccion"].ToString() != null &&
                                rows["FormaCobertura"].ToString().Length != 0 && rows["FormaCobertura"].ToString() != null &&
                                rows["CentroCosto"].ToString().Length != 0 && rows["CentroCosto"].ToString() != null &&
                                rows["NumPlaza"].ToString().Length != 0 && rows["NumPlaza"].ToString() != null &&
                                rows["CargoProveer"].ToString().Length != 0 && rows["CargoProveer"].ToString() != null &&
                                rows["MotivoSolicitud"].ToString().Length != 0 && rows["MotivoSolicitud"].ToString() != null &&
                                rows["Contrato"].ToString().Length != 0 && rows["Contrato"].ToString() != null &&
                                rows["TipoContrato"].ToString().Length != 0 && rows["TipoContrato"].ToString() != null &&

                               rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["PublicarSueldo"].ToString().Length != 0 && rows["PublicarSueldo"].ToString() != null &&
                                rows["ListaRequisiciones"].ToString().Length != 0 && rows["ListaRequisiciones"].ToString() != null &&
                                rows["CODGRSE"].ToString().Length != 0 && rows["CODGRSE"].ToString() != null &&
                                rows["CodEmpresa"].ToString().Length != 0 && rows["CodEmpresa"].ToString() != null &&
                                rows["CargoSolicitado"].ToString().Length != 0 && rows["CargoSolicitado"].ToString() != null &&
                                rows["CargoAprueba"].ToString().Length != 0 && rows["CargoAprueba"].ToString() != null &&
                                rows["CodCenCos"].ToString().Length != 0 && rows["CodCenCos"].ToString() != null &&
                                rows["EmpresaResponsable"].ToString().Length != 0 && rows["EmpresaResponsable"].ToString() != null &&
                                // Datos Aprobador No.1
                                rows["AprobadorUser1"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["AprobadorPass1"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null


                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string CodPagina = rows["CodPagina"].ToString();
                                string url = rows["url"].ToString();
                                string GrupoRequisiciones = rows["GrupoRequisiciones"].ToString();
                                string FormaCobertura = rows["FormaCobertura"].ToString();
                                string FiltroSeleccion = rows["FiltroSeleccion"].ToString();
                                string CentroCosto = rows["CentroCosto"].ToString();
                                string NumPlaza = rows["NumPlaza"].ToString();
                                string CargoProveer = rows["CargoProveer"].ToString();
                                string MotivoSolicitud = rows["MotivoSolicitud"].ToString();
                                string Contrato = rows["Contrato"].ToString();
                                string TipoContrato = rows["TipoContrato"].ToString();

                                string Ciudad = rows["Ciudad"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string PublicarSueldo = rows["PublicarSueldo"].ToString();
                                string ListaRequisiciones = rows["ListaRequisiciones"].ToString();
                                string CODGRSE = rows["CODGRSE"].ToString();
                                string CodEmpresa = rows["CodEmpresa"].ToString();
                                string CargoSolicitado = rows["CargoSolicitado"].ToString();
                                string CargoAprueba = rows["CargoAprueba"].ToString();
                                string CodCenCos = rows["CodCenCos"].ToString();
                                string EmpresaResponsable = rows["EmpresaResponsable"].ToString();
                                string AprobadorUser1 = rows["AprobadorUser1"].ToString();
                                string AprobadorPass1 = rows["AprobadorPass1"].ToString();
                                string Empresa = rows["Empresa"].ToString();
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


                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, user);

                                    string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{AprobadorUser1}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud1, database, user);


                                    ////Comienzo de Prueba
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(@id,'pLider')]");
                                        Thread.Sleep(2000);
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[2]/a");
                                        selenium.Screenshot("Selección de Personal", true, file);
                                        Thread.Sleep(200);
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[2]/ul/li/a");
                                        selenium.Screenshot("Requisición de Personal", true, file);
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodEmpr']", Empresa);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Empresa", true, file);
                                    }
                                    else
                                    {
                                        selenium.Click("//button[contains(.,'LIDER')]");
                                        Thread.Sleep(2000);
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[3]/a");
                                        selenium.Screenshot("Selección de Personal", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Requisición de Personal", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Empresa", true, file);
                                    }

                                    //CLIC NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    selenium.Screenshot("Requisición de Personal Nuevo", true, file);
                                    Thread.Sleep(2000);
                                    //GRUPO REQUI
                                    selenium.Scroll("//select[contains(@id,'ddlCodGrse')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequisiciones);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Grupo", true, file);
                                    //FORMA COBERTURA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", FormaCobertura);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cobertura", true, file);
                                    //FILTRO SELECCION
                                    selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FiltroSeleccion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Filro Selección", true, file);
                                    //CENTRO COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CentroCosto);
                                    Thread.Sleep(5000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Centro costo", true, file);
                                    //CARGO
                                    selenium.Scroll("//input[contains(@id,'txtCodCarp')]");
                                    selenium.SendKeys("//input[contains(@id,'txtCodCarp')]", CargoProveer);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Cargo", true, file);
                                    //NUMERO PLAZA
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_txtNroPlaz')]");
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_txtNroPlaz')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtNroPlaz')]", NumPlaza);
                                    Thread.Sleep(2000);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Screenshot("Plazas", true, file);
                                    Thread.Sleep(1000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoSolicitud);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo Solicitud", true, file);
                                    
                                    Thread.Sleep(1500);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicarSueldo);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar Sueldo", true, file);

                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", Contrato);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlContrDeta')]", TipoContrato);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo Contrato", true, file);
                                    Thread.Sleep(1000);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]");
                                    Thread.Sleep(1500);
                                    //OBSERVACIONES
                                    selenium.SendKeys("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    //DETALLE
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ_txtTexto')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ_txtTexto')]", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    //CIUDAD
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(500);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Screenshot("Ciudad", true, file);
                                    Thread.Sleep(1000);

                                    
                                    //GUARDAR DATOS
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(6000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(10000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(1000);
                                    selenium.Scroll("//td[8]/a/i");
                                    selenium.Screenshot("Registro exitoso", true, file);
                                    Thread.Sleep(3000);
                                    selenium.Close();

                                    ////Aprobador 
                                    selenium.LoginApps(app, AprobadorUser1, AprobadorPass1, url, file);
                                    if (database == "SQL")
                                    {
                                        selenium.Click("//span[contains(@id,'pLider')]");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Rol Lider", true, file);
                                        Thread.Sleep(500);
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[2]/a");
                                        Thread.Sleep(500);
                                        selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[2]/ul/li[2]/a");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Requisiciones por aprobar", true, file);
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodEmpr']", Empresa);
                                        Thread.Sleep(2000);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Empresa", true, file);
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(@id,'pLider')]");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Rol Lider", true, file);
                                        Thread.Sleep(500);
                                        selenium.Click("//a[contains(.,'SELECCION PERSONAL')]");
                                        Thread.Sleep(500);
                                        selenium.Click("//a[contains(.,'Aprobacion de Requisiciones')]");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Requisiciones por aprobar", true, file);
                                        Thread.Sleep(5000);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Empresa", true, file);
                                    }

                                    Thread.Sleep(1000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblTipOrd_1']");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(1000);

                                    if (selenium.ExistControl("//td[8]/a/i"))
                                    {

                                        selenium.Click("//td[8]/a/i");
                                        Thread.Sleep(2000);

                                        selenium.Screenshot("Elegir Requisición", true, file);

                                        Thread.Sleep(200);
                                        selenium.Scroll("//div[@id='printable']");
                                        selenium.Click("//td[6]/a/i");
                                        Thread.Sleep(200);
                                        selenium.Screenshot("Detalle", true, file);
                                        Thread.Sleep(200);
                                        selenium.Scroll("//input[contains(@id,'Aprueba')]");
                                        Thread.Sleep(200);
                                        selenium.Click("//input[contains(@id,'Aprueba')]");
                                        Thread.Sleep(500);
                                        Thread.Sleep(200);
                                        selenium.AcceptAlert();
                                        selenium.Screenshot("Aprueba Requisición", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Envia correo Requisición", true, file);
                                    }
                                    else
                                    {
                                        Assert.Fail("ERROR: NO EXISTEN REQUISICIONES POR APROBAR");
                                    }


                                    Thread.Sleep(5000);
                                    selenium.Close();
                                    Thread.Sleep(5000);
                                    fv.ConvertWordToPDF(file, database);
                                    string eliminarSolicitud21 = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud21, database, user);

                                    string eliminarRequi1 = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi1, database, user);

                                    string eliminarSolicitud11 = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{AprobadorUser1}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud11, database, user);
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
                                    Thread.Sleep(3000);
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

        public void SL_FlujoAprobaciónRequisicionesCargoSolicitado()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_FlujoAprobaciónRequisicionesCargoSolicitado")
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
                                //Datos Login Requisición Personal 
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                 rows["AprobadorUser1"].ToString().Length != 0 && rows["AprobadorUser1"].ToString() != null &&
                                rows["AprobadorPass1"].ToString().Length != 0 && rows["AprobadorPass1"].ToString() != null &&
                                 rows["AprobadorUser2"].ToString().Length != 0 && rows["AprobadorUser2"].ToString() != null &&
                                rows["AprobadorPass2"].ToString().Length != 0 && rows["AprobadorPass2"].ToString() != null &&

                                //Datos Requisición Personal 
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["GrupoRequisiciones"].ToString().Length != 0 && rows["GrupoRequisiciones"].ToString() != null &&
                                rows["FormaCobertura"].ToString().Length != 0 && rows["FormaCobertura"].ToString() != null &&
                                rows["FiltroSeleccion"].ToString().Length != 0 && rows["FiltroSeleccion"].ToString() != null &&
                                rows["CentroCosto"].ToString().Length != 0 && rows["CentroCosto"].ToString() != null &&
                                rows["NumPlaza"].ToString().Length != 0 && rows["NumPlaza"].ToString() != null &&
                                rows["CargoProveer"].ToString().Length != 0 && rows["CargoProveer"].ToString() != null &&
                                rows["MotivoSolicitud"].ToString().Length != 0 && rows["MotivoSolicitud"].ToString() != null &&
                                rows["Contrato"].ToString().Length != 0 && rows["Contrato"].ToString() != null &&
                                rows["TipoContrato"].ToString().Length != 0 && rows["TipoContrato"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["PublicarSueldo"].ToString().Length != 0 && rows["PublicarSueldo"].ToString() != null &&
                                rows["ListaRequisiciones"].ToString().Length != 0 && rows["ListaRequisiciones"].ToString() != null &&
                                rows["CodEmpresa"].ToString().Length != 0 && rows["CodEmpresa"].ToString() != null &&
                                rows["CodMotivo"].ToString().Length != 0 && rows["CodMotivo"].ToString() != null &&
                                rows["CargoSolicitado"].ToString().Length != 0 && rows["CargoSolicitado"].ToString() != null &&
                                rows["INDACTI"].ToString().Length != 0 && rows["INDACTI"].ToString() != null &&
                                rows["CargoAprueba"].ToString().Length != 0 && rows["CargoAprueba"].ToString() != null &&
                                rows["TipoDocumAdjunto"].ToString().Length != 0 && rows["TipoDocumAdjunto"].ToString() != null &&
                                rows["CargoEspecifico"].ToString().Length != 0 && rows["CargoEspecifico"].ToString() != null
                                // Datos Aprobador No.1

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string CodPagina = rows["CodPagina"].ToString();
                                string url = rows["url"].ToString();
                                string GrupoRequisiciones = rows["GrupoRequisiciones"].ToString();
                                string FormaCobertura = rows["FormaCobertura"].ToString();
                                string FiltroSeleccion = rows["FiltroSeleccion"].ToString();
                                string CentroCosto = rows["CentroCosto"].ToString();
                                string NumPlaza = rows["NumPlaza"].ToString();
                                string CargoProveer = rows["CargoProveer"].ToString();
                                string MotivoSolicitud = rows["MotivoSolicitud"].ToString();
                                string Contrato = rows["Contrato"].ToString();
                                string TipoContrato = rows["TipoContrato"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string PublicarSueldo = rows["PublicarSueldo"].ToString();
                                string ListaRequisiciones = rows["ListaRequisiciones"].ToString();
                                string CodEmpresa = rows["CodEmpresa"].ToString();
                                string CodMotivo = rows["CodMotivo"].ToString();
                                string CargoSolicitado = rows["CargoSolicitado"].ToString();
                                string INDACTI = rows["INDACTI"].ToString();
                                string CargoAprueba = rows["CargoAprueba"].ToString();
                                string CargoEspecifico = rows["CargoEspecifico"].ToString();
                                string AprobadorUser1 = rows["AprobadorUser1"].ToString();
                                string AprobadorPass1 = rows["AprobadorPass1"].ToString();
                                string TipoDocumAdjunto = rows["TipoDocumAdjunto"].ToString();
                                string AprobadorUser2 = rows["AprobadorUser2"].ToString();
                                string AprobadorPass2 = rows["AprobadorPass2"].ToString();
                                //limpiar procesos
                                Process[] processes = Process.GetProcessesByName("chromedriver");
                                if (processes.Length > 0)
                                {
                                    for (int i = 0; i < processes.Length; i++)
                                    {
                                        processes[i].Kill();
                                    }
                                }
                                string User;
                                string cod_empr;
                                try
                                {
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, user);

                                    string eliminarSolicitud1 = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{AprobadorUser1}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud1, database, user);

                                    string eliminarSolicitud2 = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{AprobadorUser2}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud2, database, user);


                                    ////Comienzo de Prueba
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//span[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[2]/a");
                                    selenium.Screenshot("Selección de Personal", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[2]/ul/li/a");
                                    selenium.Screenshot("Requisición de Personal", true, file);
                                    Thread.Sleep(2000);
                                    //CLIC NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    selenium.Screenshot("Requisición de Personal Nuevo", true, file);
                                    Thread.Sleep(2000);

                                    if (database == "ORA")
                                    {
                                        //GRUPO REQUISICION
                                        selenium.Scroll("//select[contains(@id,'ddlCodGrse')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequisiciones);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Grupo Requisición", true, file);
                                        //FORMA COBERTURA
                                        selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", FormaCobertura);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Forma Cobertura", true, file);
                                        //FILTRO SELECCION
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FiltroSeleccion);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Filtro Selección", true, file);
                                        //CENTRO COSTO
                                        selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                        selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CentroCosto);
                                        Thread.Sleep(5000);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(5000);
                                        SendKeys.SendWait("{ENTER}");
                                        selenium.Screenshot("Centro Costo", true, file);
                                        //NUMERO PLAZAS
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_txtNroPlaz')]");
                                        Thread.Sleep(1000);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtNroPlaz')]", NumPlaza);
                                        Thread.Sleep(2000);
                                        selenium.Click("//div[@id='ctl00_pBotones']/div");
                                        selenium.Screenshot("Numero Plazas", true, file);
                                        //CARGO
                                        selenium.Scroll("//input[contains(@id,'txtCodCarp')]");
                                        selenium.SendKeys("//input[contains(@id,'txtCodCarp')]", CargoProveer);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Cargo a Proveer", true, file);
                                        Thread.Sleep(1000);
                                        //MOTIVO SOLICITUD
                                        selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoSolicitud);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Motivo", true, file);
                                        Thread.Sleep(1500);
                                        //PUBLICAR SUELDO
                                        selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicarSueldo);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Publicar", true, file);
                                        //CONTRATO
                                        selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", Contrato);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Contrato", true, file);
                                        //TIPO DE CONTRATO
                                        selenium.Scroll("//select[contains(@id,'ctl00_ContenidoPagina_ddlContrDeta')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlContrDeta')]", TipoContrato);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Tipo Contrato", true, file);
                                        Thread.Sleep(1000);
                                        selenium.Click("//div[@id='ctl00_pBotones']/div");
                                        Thread.Sleep(1000);
                                        //OBSERVACIONES
                                        selenium.Scroll("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]");
                                        selenium.SendKeys("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]", Observacion);
                                        Thread.Sleep(1500);
                                        //DETALLE
                                        selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ_txtTexto')]");
                                        selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ_txtTexto')]", Observacion);
                                        Thread.Sleep(1500);
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(1000);
                                        //CIUDAD
                                        selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(1500);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Datos Requisición ", true, file);
                                        selenium.Click("//div[@id='ctl00_pBotones']/div");
                                        selenium.Screenshot("Ingreso de datos", true, file);
                                        Thread.Sleep(1000);
                                    }
                                    else
                                    {
                                        //GRUPO REQUISICION
                                        selenium.Scroll("//select[contains(@id,'ddlCodGrse')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequisiciones);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Grupo Requisición", true, file);
                                        //FORMA COBERTURA
                                        selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", FormaCobertura);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Cobertura", true, file);
                                        //FILTRO SELECCION
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FiltroSeleccion);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Filtro", true, file);
                                        //CENTRO COSTO
                                        selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                        selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CentroCosto);
                                        Thread.Sleep(5000);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Centro Costo", true, file);
                                        //NUMERO PLAZAS
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_txtNroPlaz')]");
                                        selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_txtNroPlaz')]");
                                        Thread.Sleep(1000);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtNroPlaz')]", NumPlaza);
                                        Thread.Sleep(2000);
                                        selenium.Click("//div[@id='ctl00_pBotones']/div");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Plazas", true, file);
                                        //CARGO
                                        selenium.Scroll("//input[contains(@id,'txtCodCarp')]");
                                        selenium.SendKeys("//input[contains(@id,'txtCodCarp')]", CargoProveer);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Cargo", true, file);
                                        Thread.Sleep(1000);
                                        //MOTIVO SOLICITUD
                                        selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoSolicitud);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Motivo", true, file);
                                        selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                        Thread.Sleep(1500);
                                        //PUBLICAR SUELDO
                                        selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicarSueldo);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Publicar", true, file);
                                        //CONTRATO
                                        selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", Contrato);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Contrato", true, file);
                                        //TIPO DE CONTRATO
                                        selenium.Scroll("//select[contains(@id,'ctl00_ContenidoPagina_ddlContrDeta')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlContrDeta')]", TipoContrato);////select[@id='ctl00_ContenidoPagina_ddlContrDeta']
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Ingreso de datos", true, file);
                                        Thread.Sleep(1000);
                                        selenium.Click("//div[@id='ctl00_pBotones']/div");
                                        selenium.Scroll("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]");
                                        Thread.Sleep(1000);
                                        //OBSERVACIONES
                                        selenium.Scroll("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]");
                                        selenium.SendKeys("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]", Observacion);
                                        Thread.Sleep(1500);
                                        //DETALLE
                                        selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ_txtTexto')]");
                                        selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ_txtTexto')]", Observacion);
                                        Thread.Sleep(1500);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Observaciones", true, file);
                                        //CIUDAD
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(1000);
                                        selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(1500);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                        Thread.Sleep(1500);
                                        selenium.Click("//div[@id='ctl00_pBotones']/div");
                                        selenium.Screenshot("Ingreso de datos", true, file);
                                        Thread.Sleep(1000);

                                    }
                                    //GUARDAR DATOS
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(5000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(5000);
                                    selenium.Scroll("//td[8]/a/i");
                                    selenium.Screenshot("Registro exitoso", true, file);
                                    selenium.Close();


                                    //Aprobador 1
                                    selenium.LoginApps(app, AprobadorUser1, AprobadorPass1, url, file);
                                    Thread.Sleep(1000);
                                    selenium.Click("//span[contains(@id,'pLider')]");
                                    Thread.Sleep(1000);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Aprobador Rol Lider", true, file);
                                    Thread.Sleep(500);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[2]/a");
                                    Thread.Sleep(500);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[2]/ul/li[2]/a");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblTipOrd_1']");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(1000);

                                    if (selenium.ExistControl("//td[8]/a/i"))
                                    {

                                        Thread.Sleep(5000);
                                        selenium.Click("//td[8]/a/i");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Elegir Requisición", true, file);
                                        Thread.Sleep(200);
                                        selenium.Scroll("//div[@id='printable']");
                                        selenium.Click("//td[6]/a/i");
                                        Thread.Sleep(200);
                                        selenium.Screenshot("Cargo Solicitado Requisición", true, file);
                                        Thread.Sleep(200);
                                        selenium.Scroll("//input[contains(@id,'Aprueba')]");
                                        selenium.Click("//input[contains(@id,'Aprueba')]");
                                        Thread.Sleep(500);
                                        Thread.Sleep(200);
                                        selenium.AcceptAlert();
                                        selenium.Screenshot("Aprueba Requisición", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Envia correo Requisición", true, file);
                                    }
                                    else
                                    {
                                        Assert.Fail("ERROR: NO EXISTEN REQUISICIONES POR APROBAR");
                                    }

                                    //ENVIO CORREO
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Envia Correo Aprobación", true, file);
                                    selenium.Click("//input[contains(@id,'btnEnviar')]");
                                    Thread.Sleep(3000);
                                    selenium.Close();

                                    //Aprobador 2
                                    selenium.LoginApps(app, AprobadorUser2, AprobadorPass2, url, file);
                                    Thread.Sleep(1000);
                                    selenium.Click("//span[contains(@id,'pLider')]");
                                    Thread.Sleep(1000);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Aprobador Rol Lider", true, file);
                                    Thread.Sleep(500);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[2]/a");
                                    Thread.Sleep(500);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[2]/ul/li[2]/a");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblTipOrd_1']");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(1000);

                                    if (selenium.ExistControl("//td[8]/a/i"))
                                    {

                                        Thread.Sleep(5000);
                                        selenium.Click("//td[8]/a/i");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Elegir Requisición", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Scroll("//div[@id='printable']");
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_Estsoli_ctl03_LinkButton1']/i");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Cargo Solicitado Requisición", true, file);
                                        Thread.Sleep(200);
                                        selenium.Scroll("//input[contains(@id,'Aprueba')]");
                                        selenium.Click("//input[contains(@id,'Aprueba')]");
                                        Thread.Sleep(2000);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Aprueba Requisición", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Envia correo Requisición", true, file);
                                    }
                                    else
                                    {
                                        Assert.Fail("ERROR: NO EXISTEN REQUISICIONES POR APROBAR");
                                    }

                                    //ENVIO CORREO
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Envia Correo Aprobación", true, file);
                                    selenium.Click("//input[contains(@id,'btnEnviar')]");
                                    Thread.Sleep(3000);
                                    selenium.Close();

                                    //VERIFICAR
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//span[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[2]/a");
                                    selenium.Screenshot("Selección de Personal", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[2]/ul/li/a");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Requisición de Personal", true, file);
                                    //DETALLE
                                    selenium.Scroll("//*[@id='tablaDatos']/tbody/tr/td[7]/a");
                                    Thread.Sleep(3000);
                                    selenium.Click("//*[@id='tablaDatos']/tbody/tr/td[7]/a");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Requisición Aprobada", true, file);
                                    Thread.Sleep(3000);
                                    selenium.Close();
                                    Thread.Sleep(3000);

                                    string eliminarSolicitud11 = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud11, database, user);

                                    string eliminarRequi1 = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi1, database, user);

                                    string eliminarSolicitud22 = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{AprobadorUser1}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud22, database, user);

                                    string eliminarSolicitu32 = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{AprobadorUser2}'";
                                    db.UpdateDeleteInsert(eliminarSolicitu32, database, user);
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
        public void SL_RequisiciónDePersonal()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_RequisiciónDePersonal")
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
                                rows["Contenido"].ToString().Length != 0 && rows["Contenido"].ToString() != null &&
                                rows["CodPagina"].ToString().Length != 0 && rows["CodPagina"].ToString() != null &&
                                rows["EmailRecepcion"].ToString().Length != 0 && rows["EmailRecepcion"].ToString() != null &&
                                rows["PassEmailRecepcion"].ToString().Length != 0 && rows["PassEmailRecepcion"].ToString() != null &&
                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["Requisiciones"].ToString().Length != 0 && rows["Requisiciones"].ToString() != null &&
                                rows["Vacante"].ToString().Length != 0 && rows["Vacante"].ToString() != null &&
                                rows["Filtro"].ToString().Length != 0 && rows["Filtro"].ToString() != null &&
                                rows["Costo"].ToString().Length != 0 && rows["Costo"].ToString() != null &&
                                rows["Plazas"].ToString().Length != 0 && rows["Plazas"].ToString() != null &&
                                rows["Motivo"].ToString().Length != 0 && rows["Motivo"].ToString() != null &&
                                rows["Contrato"].ToString().Length != 0 && rows["Contrato"].ToString() != null &&
                                rows["TipoContrato"].ToString().Length != 0 && rows["TipoContrato"].ToString() != null &&
                                rows["Sueldo"].ToString().Length != 0 && rows["Sueldo"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["Publicar"].ToString().Length != 0 && rows["Publicar"].ToString() != null)
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string CodPagina = rows["CodPagina"].ToString();
                                string EmailRecepcion = rows["EmailRecepcion"].ToString();
                                string PassEmailRecepcion = rows["PassEmailRecepcion"].ToString();
                                string Descripcion = rows["Descripcion"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Contenido = rows["Contenido"].ToString();
                                string Requisiciones = rows["Requisiciones"].ToString();
                                string Vacante = rows["Vacante"].ToString();
                                string Filtro = rows["Filtro"].ToString();
                                string Costo = rows["Costo"].ToString();
                                string Plazas = rows["Plazas"].ToString();
                                string Motivo = rows["Motivo"].ToString();
                                string Contrato = rows["Contrato"].ToString();
                                string TipoContrato = rows["TipoContrato"].ToString();
                                string Sueldo = rows["Sueldo"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string Publicar = rows["Publicar"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, user);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, user);

                                    //INICIO
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "ORA")
                                    {
                                        selenium.Click("//span[contains(.,'Lider')]");
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                    }

                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Seleccionar Rol Lider", true, file);

                                    if (database == "ORA")
                                    {
                                        //INGRESO REQUISICION
                                        Thread.Sleep(2000);
                                        selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                        Thread.Sleep(1500);
                                        //REQUISICION PERSONAL
                                        selenium.Scroll("//a[contains(.,'Requisición de Personal')]");
                                        selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Requisición de Personal", true, file);
                                        Thread.Sleep(2000);
                                        //NUEVO
                                        selenium.Click("//a[@id='ctl00_btnNuevo']");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Nuevo", true, file);
                                        Thread.Sleep(2000);
                                        //GRUPO REQUISICIONES
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodGrse']", Requisiciones);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Requisiciones", true, file);
                                        //VACANTE
                                        selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlForCobe']");
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlForCobe']", Vacante);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Vacante", true, file);
                                        //FILTRO
                                        selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlFilSele']");
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlFilSele']", Filtro);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                        //COSTO
                                        selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                        selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", Costo);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(2000);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Costo", true, file);
                                        //PLAZAS
                                        selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtNroPlaz']");
                                        selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNroPlaz']", Plazas);
                                        selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNroPlaz']", Plazas);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Plazas", true, file);
                                        //CARGO
                                        selenium.Scroll("//input[contains(@id,'txtCodCarp')]");
                                        selenium.SendKeys("//input[contains(@id,'txtCodCarp')]", "102006");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Cargo", true, file);
                                        //MOTIVO
                                        selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlCodMoti']");
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodMoti']", Motivo);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Motivo", true, file);
                                        //CONTRATO
                                        selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlTipCont']");
                                        Thread.Sleep(2000);
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlTipCont']", Contrato);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Contrato", true, file);
                                        //TIPO CONTRATO
                                        selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlContrDeta']");
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlContrDeta']", TipoContrato);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Tipo contrato", true, file);
                                        //PUBLICAR
                                        selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlVisSuew']");
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlVisSuew']", Publicar);
                                        selenium.Screenshot("Publicar", true, file);
                                        //DETALLE
                                        selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ_txtTexto')]");
                                        selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ_txtTexto')]", "PRUEBAS");
                                        Thread.Sleep(500);
                                        Thread.Sleep(2000);
                                        selenium.Click("//div[@id='ctl00_pBotones']/div");
                                        selenium.Scroll("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]");
                                        Thread.Sleep(2000);
                                        //OBSERVACIONES
                                        selenium.SendKeys("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]", "PRUEBAS");
                                        Thread.Sleep(500);
                                        //Validación 1 País
                                        Thread.Sleep(500);
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(500);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Ingreso de País para validar", true, file);
                                        Thread.Sleep(500);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Validación de campo en blanco", true, file);
                                        string Texto1 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        if (Texto1 != "")
                                        {
                                            errorMessages.Add(" ::::::::::::::::::::::" + "MSG: No se eliminó el contenido del campo Ciudad al hacer TAB, el texto encontado es: " + Texto1);
                                        }

                                        //Validación 2 Departamento
                                        Thread.Sleep(500);
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(500);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                        Thread.Sleep(500);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Validación de campo en blanco", true, file);

                                        string Texto2 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        if (Texto2 != "")
                                        {
                                            errorMessages.Add(" ::::::::::::::::::::::" + "MSG: No se eliminó el contenido del campo Ciudad al hacer TAB, el texto encontado es: " + Texto2);
                                        }

                                        //Validación Caracteres especiales
                                        Thread.Sleep(500);
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(500);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                        Thread.Sleep(500);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Validación de campo en blanco", true, file);


                                        string Texto3 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        if (Texto3 != "")
                                        {
                                            errorMessages.Add(" ::::::::::::::::::::::" + "MSG: No se eliminó el contenido del campo Ciudad al hacer TAB, el texto encontado es: " + Texto3);
                                        }

                                        //Validación Exitosa
                                        Thread.Sleep(500);
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(500);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                        Thread.Sleep(500);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Validación de campo en blanco", true, file);


                                        string Texto4 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        if (Texto4 == "")
                                        {
                                            errorMessages.Add(" ::::::::::::::::::::::" + "MSG: Se eliminó el contenido del campo Ciudad al hacer TAB");
                                        }
                                        //NUMERO PLAZAS
                                        selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtNroPlaz']");
                                        Thread.Sleep(1000);
                                        selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNroPlaz']", Plazas);
                                        selenium.Screenshot("Datos Ingresados", true, file);
                                        Thread.Sleep(2000);
                                        //FUNCIONARIO A REEMPLAZAR
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(6000);
                                        selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(1000);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(6000);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(6000);
                                        selenium.SendKeys("//input[@type='search']", Funcionario);
                                        selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                        Thread.Sleep(6000);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(6000);

                                        //ADJUNTO
                                        Thread.Sleep(1000);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(2000);
                                        SendKeys.SendWait(Ruta);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(1000);
                                        selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Archivo adjunto", true, file);
                                        //GUARDAR
                                        selenium.Scroll("//a[@id='btnGuardar']");
                                        selenium.Click("//a[@id='btnGuardar']");
                                        Thread.Sleep(1000);
                                        Thread.Sleep(2000);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(2000);
                                        Thread.Sleep(2000);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(10000);
                                        //VERIFICACION REGISTRO
                                        selenium.Scroll("//td[8]/a/i");
                                        selenium.Screenshot("Registro exitoso", true, file);
                                        Thread.Sleep(5000);
                                    }
                                    else
                                    {
                                        //INGRESO REQUISICION
                                        Thread.Sleep(2000);
                                        selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                        Thread.Sleep(1500);
                                        //REQUISICION PERSONAL
                                        selenium.Scroll("//a[contains(.,'Requisición de Personal')]");
                                        selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Requisición de Personal", true, file);
                                        Thread.Sleep(2000);
                                        //NUEVO
                                        selenium.Click("//a[@id='ctl00_btnNuevo']");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Nuevo", true, file);
                                        Thread.Sleep(2000);
                                        //GRUPO REQUISICIONES
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodGrse']", Requisiciones);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Requisiciones", true, file);
                                        //VACANTE
                                        selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlForCobe']");
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlForCobe']", Vacante);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Vacante", true, file);
                                        //FILTRO
                                        selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlFilSele']");
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlFilSele']", Filtro);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                        //COSTO
                                        selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                        selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", Costo);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(2000);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Costo", true, file);
                                        //PLAZAS
                                        selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtNroPlaz']");
                                        selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNroPlaz']", Plazas);
                                        selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNroPlaz']", Plazas);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Plazas", true, file);
                                        //CARGO
                                        selenium.Scroll("//input[contains(@id,'txtCodCarp')]");
                                        selenium.SendKeys("//input[contains(@id,'txtCodCarp')]", "001020");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Cargo", true, file);
                                        //MOTIVO
                                        selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlCodMoti']");
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodMoti']", Motivo);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Motivo", true, file);
                                        //CONTRATO
                                        selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlTipCont']");
                                        Thread.Sleep(2000);
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlTipCont']", Contrato);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Contrato", true, file);
                                        //TIPO CONTRATO
                                        selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlContrDeta']");
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlContrDeta']", TipoContrato);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Tipo contrato", true, file);
                                        //PUBLICAR
                                        selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlVisSuew']");
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlVisSuew']", Publicar);
                                        selenium.Screenshot("Publicar", true, file);
                                        //DETALLE
                                        selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ_txtTexto')]");
                                        selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ_txtTexto')]", "PRUEBAS");
                                        Thread.Sleep(500);
                                        Thread.Sleep(2000);
                                        selenium.Click("//div[@id='ctl00_pBotones']/div");
                                        selenium.Scroll("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]");
                                        Thread.Sleep(2000);
                                        //OBSERVACIONES
                                        selenium.SendKeys("//textarea[contains(@id,'ctl00_ContenidoPagina_KtxtObserSoli_txtTexto')]", "PRUEBAS");
                                        Thread.Sleep(500);
                                        //Validación 1 País
                                        Thread.Sleep(500);
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(500);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Ingreso de País para validar", true, file);
                                        Thread.Sleep(500);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Validación de campo en blanco", true, file);
                                        string Texto1 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        if (Texto1 != "")
                                        {
                                            errorMessages.Add(" ::::::::::::::::::::::" + "MSG: No se eliminó el contenido del campo Ciudad al hacer TAB, el texto encontado es: " + Texto1);
                                        }

                                        //Validación 2 Departamento
                                        Thread.Sleep(500);
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]"); 
                                        Thread.Sleep(500);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                        Thread.Sleep(500);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Validación de campo en blanco", true, file);

                                        string Texto2 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        if (Texto2 != "")
                                        {
                                            errorMessages.Add(" ::::::::::::::::::::::" + "MSG: No se eliminó el contenido del campo Ciudad al hacer TAB, el texto encontado es: " + Texto2);
                                        }

                                        //Validación Caracteres especiales
                                        Thread.Sleep(500);
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(500);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                        Thread.Sleep(500);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Validación de campo en blanco", true, file);


                                        string Texto3 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        if (Texto3 != "")
                                        {
                                            errorMessages.Add(" ::::::::::::::::::::::" + "MSG: No se eliminó el contenido del campo Ciudad al hacer TAB, el texto encontado es: " + Texto3);
                                        }

                                        //Validación Exitosa
                                        Thread.Sleep(500);
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(500);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                        Thread.Sleep(500);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Validación de campo en blanco", true, file);


                                        string Texto4 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        if (Texto4 == "")
                                        {
                                            errorMessages.Add(" ::::::::::::::::::::::" + "MSG: Se eliminó el contenido del campo Ciudad al hacer TAB");
                                        }
                                        //NUMERO PLAZAS
                                        selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtNroPlaz']");
                                        Thread.Sleep(1000);
                                        selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNroPlaz']", Plazas);
                                        selenium.Screenshot("Datos Ingresados", true, file);
                                        Thread.Sleep(2000);
                                        //FUNCIONARIO A REEMPLAZAR
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(6000);
                                        selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                        Thread.Sleep(1000);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(6000);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(6000);
                                        selenium.SendKeys("//input[@type='search']", Funcionario);
                                        selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                        Thread.Sleep(6000);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(6000);

                                        //ADJUNTO
                                        Thread.Sleep(1000);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(2000);
                                        SendKeys.SendWait(Ruta);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(1000);
                                        selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Archivo adjunto", true, file);
                                        //GUARDAR
                                        selenium.Scroll("//a[@id='btnGuardar']");
                                        selenium.Click("//a[@id='btnGuardar']");
                                        Thread.Sleep(1000);
                                        Thread.Sleep(2000);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(2000);
                                        Thread.Sleep(2000);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(10000);
                                        //VERIFICACION REGISTRO
                                        selenium.Scroll("//td[8]/a/i");
                                        selenium.Screenshot("Registro exitoso", true, file);
                                        Thread.Sleep(5000);
                                    }
                                    Thread.Sleep(2000);
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
        public void SL_AprobaciónRequisiciónDePersonalJefeInmediato()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_AprobaciónRequisiciónDePersonalJefeInmediato")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string JefeUser = rows["JefeUser"].ToString();
                                string JefePass = rows["JefePass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();

                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }

                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    ////Process:Login//////////////////////////////////
                                    string error = string.Empty;
                                    string errorfilter = string.Empty;
                                    List<string> errorMessagesMetodo = new List<string>();

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);

                                    //Process: Requisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Requisición de Personal", true, file);
                                    //NUEVO
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Nueva Requisición de Personal", true, file);
                                    //GRUPO
                                    selenium.SelectElementByName("//select[contains(@id,'ContenidoPagina_ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(1500);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Requisición de Personal", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ContenidoPagina_ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ContenidoPagina_ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(1000);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Requisición de Personal", true, file);

                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ctl00_ContenidoPagina_ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(1000);
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(1500);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Costo", true, file);
                                    //PLAZAS
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(2000); selenium.Click("//input[contains(@id,'txtNroPlaz')]");
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Plazas", true, file);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ContenidoPagina_ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ContenidoPagina_ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(1500);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Cargo", true, file);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ContenidoPagina_ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ContenidoPagina_ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(1500);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Motivo", true, file);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ContenidoPagina_ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ContenidoPagina_ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(1500);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Contrato", true, file);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_ddlContrDeta']");
                                    selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_ddlContrDeta']", TipoContratoRequi);
                                    Thread.Sleep(1500);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //PUBLICAR
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_ddlVisSuew']");
                                    selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_ddlVisSuew']", PublicardadRequi);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Publicar", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ciudad", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1500);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ_txtTexto')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ_txtTexto')]", ComentarioRequi);
                                    Thread.Sleep(1500);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Comentarios", true, file);
                                    //PLAZAS
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(1500);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Plazas", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(1500);
                                    Thread.Sleep(10000);
                                    selenium.AcceptAlert();
                                    selenium.AcceptAlert();
                                    Thread.Sleep(10000);
                                    selenium.Scroll("//td[8]/a/i");
                                    selenium.Screenshot("Registro exitoso", true, file);
                                    Thread.Sleep(1500);
                                    selenium.Close();

                                    //LIDER
                                    selenium.LoginApps(app, JefeUser, JefePass, url, file);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Aprobación de Requisiciones')]");
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Aprobación de Requisiciones", true, file);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_rblTipOrd_1']");
                                    Thread.Sleep(1000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(1000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_soldap_ctl03_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle", true, file);
                                    Thread.Sleep(1000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_Estsoli_ctl02_LinkButton1']/i");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Aprobación", true, file);

                                    if (selenium.ExistControl("//input[contains(@id,'Aprueba')]"))
                                    {
                                        selenium.Scroll("//input[contains(@id,'Aprueba')]");
                                        Thread.Sleep(500);
                                        Thread.Sleep(500);
                                        selenium.Click("//input[contains(@id,'Aprueba')]");
                                        Thread.Sleep(500);


                                        Thread.Sleep(1000);
                                        selenium.AcceptAlert();

                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Aprobación terminada", true, file);
                                    }
                                    else
                                    {
                                        selenium.Screenshot("Sin boton aprobación", true, file);
                                    }
                                    //ENVIO CORREO
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Envia Correo Aprobación", true, file);
                                    selenium.Click("//input[contains(@id,'btnEnviar')]");
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    //////
                                    fv.ConvertWordToPDF(file, database);
                                    string eliminarSolicitud11 = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud11, database, User);

                                    string eliminarRequi22 = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi22, database, User);

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
        public void SL_CancelaciónRequisiciónDespuésSolicitudConvocatoriaAbierta()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_CancelaciónRequisiciónDespuésSolicitudConvocatoriaAbierta")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Rquisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(2000);
                                    //PLAZAS
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    selenium.Click("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Plazas", true, file);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //ADJUNTO
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
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
        public void SL_CancelaciónRequisiciónDespuésSolicitudConvocatoriaCerrada()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_CancelaciónRequisiciónDespuésSolicitudConvocatoriaCerrada")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Rquisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Grupo Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(2000);
                                    //PLAZAS
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    selenium.Click("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Plazas", true, file);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //ADJUNTO
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database);
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
        public void SL_CancelaciónRequisiciónDespuésSolicitudConvocatoriaMixta()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_CancelaciónRequisiciónDespuésSolicitudConvocatoriaMixta")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Rquisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Grupo Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(2000);
                                    //PLAZAS
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    selenium.Click("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Plazas", true, file);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //ADJUNTO
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(2000);
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
        public void SL_CancelaciónRequisiciónDespuésSolicitudConvocatoriaDirecta()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_CancelaciónRequisiciónDespuésSolicitudConvocatoriaDirecta")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string PersonalPropuesto = rows["PersonalPropuesto"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Rquisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Grupo Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(2000);
                                    //PLAZAS
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    selenium.Click("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Plazas", true, file);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //BUSQUEDA PERSONAL PROPUESTO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']", PersonalPropuesto);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    Thread.Sleep(6000);
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleadosD_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleadosD_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Personal Propuesto", true, file);
                                    Thread.Sleep(2000);
                                    //ADJUNTO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    Thread.Sleep(1000);
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database);
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
        public void SL_CancelaciónRequisiciónDespuésSolicitudPromoción()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_CancelaciónRequisiciónDespuésSolicitudPromoción")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string PersonalPropuesto = rows["PersonalPropuesto"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Rquisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisicion", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(2000);
                                    //PLAZAS
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    selenium.Click("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Plazas", true, file);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //BUSQUEDA PERSONAL PROPUESTO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']", PersonalPropuesto);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    Thread.Sleep(6000);
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Personal Porpuesto", true, file);
                                    Thread.Sleep(2000);
                                    //ADJUNTO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    Thread.Sleep(1000);
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database);
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
        public void SL_CancelaciónRequisiciónDespuésSolicitudTransferencias()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_CancelaciónRequisiciónDespuésSolicitudTransferencias")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string PersonalPropuesto = rows["PersonalPropuesto"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Rquisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.Scroll("//select[contains(@id,'ddlCodGrse')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Grupo Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(2000);
                                    //PLAZAS
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    selenium.Click("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Plazas", true, file);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //BUSQUEDA PERSONAL PROPUESTO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']", PersonalPropuesto);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    Thread.Sleep(6000);
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Personal Porpuesto", true, file);
                                    Thread.Sleep(2000);
                                    //ADJUNTO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    Thread.Sleep(1000);
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro requisicion", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    fv.ConvertWordToPDF(file, database);
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
        public void SL_SolicitudRequisiciónMultiempresaConvocatoriaAbiertaContratoFijo()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_SolicitudRequisiciónMultiempresaConvocatoriaAbiertaContratoFijo")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string Empresa = rows["Empresa"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Requisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", Empresa);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                   
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    Thread.Sleep(3000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    Thread.Sleep(3000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //ADJUNTO
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro requisicion", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", Empresa);
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(6000);
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
        public void SL_SolicitudRequisiciónMultiempresaConvocatoriaAbiertaContratoIndefinido()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_SolicitudRequisiciónMultiempresaConvocatoriaAbiertaContratoIndefinido")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string ddlCodEmprs = rows["ddlCodEmprs"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Requisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.Scroll("//select[contains(@id,'ddlCodGrse')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //ADJUNTO
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(6000);
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
        public void SL_SolicitudRequisiciónMultiempresaConvocatoriaCerradaContratoFijo()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_SolicitudRequisiciónMultiempresaConvocatoriaCerradaContratoFijo")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string ddlCodEmprs = rows["ddlCodEmprs"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Requisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //ADJUNTO
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(6000);
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
        public void SL_SolicitudRequisiciónMultiempresaConvocatoriaCerradaContratoIndefinido()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_SolicitudRequisiciónMultiempresaConvocatoriaCerradaContratoIndefinido")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string ddlCodEmprs = rows["ddlCodEmprs"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Requisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //ADJUNTO
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(6000);
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
        public void SL_SolicitudRequisiciónMultiempresaConvocatoriaMixtaContratoFijo()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_SolicitudRequisiciónMultiempresaConvocatoriaMixtaContratoFijo")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string ddlCodEmprs = rows["ddlCodEmprs"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Requisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //ADJUNTO
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(6000);
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
        public void SL_SolicitudRequisiciónMultiempresaConvocatoriaMixtaContratoIndefinido()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_SolicitudRequisiciónMultiempresaConvocatoriaMixtaContratoIndefinido")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string ddlCodEmprs = rows["ddlCodEmprs"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Requisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //ADJUNTO
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(6000);
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
        public void SL_SolicitudRequisiciónMultiempresaPromociónContratoFijo()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_SolicitudRequisiciónMultiempresaPromociónContratoFijo")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string PersonalPropuesto = rows["PersonalPropuesto"].ToString();
                                string ddlCodEmprs = rows["ddlCodEmprs"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Rquisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //BUSQUEDA PERSONAL PROPUESTO
                                    selenium.Scroll("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']", PersonalPropuesto);
                                    Thread.Sleep(2000);
                                   // selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    Thread.Sleep(6000);
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Personal Porpuesto", true, file);
                                    Thread.Sleep(2000);
                                    //ADJUNTO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    Thread.Sleep(1000);
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(6000);
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
        public void SL_SolicitudRequisiciónMultiempresaPromociónContratoIndefinido()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_SolicitudRequisiciónMultiempresaPromociónContratoIndefinido")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string PersonalPropuesto = rows["PersonalPropuesto"].ToString();
                                string ddlCodEmprs = rows["ddlCodEmprs"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Rquisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //BUSQUEDA PERSONAL PROPUESTO
                                    selenium.Scroll("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']", PersonalPropuesto);
                                    Thread.Sleep(2000);
                                    //selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    Thread.Sleep(6000);
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Personal Porpuesto", true, file);
                                    Thread.Sleep(2000);
                                    //ADJUNTO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    Thread.Sleep(1000);
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(6000);
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
        public void SL_SolicitudRequisiciónMultiempresaTransferenciaContratoFijo()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_SolicitudRequisiciónMultiempresaTransferenciaContratoFijo")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string PersonalPropuesto = rows["PersonalPropuesto"].ToString();
                                string ddlCodEmprs = rows["ddlCodEmprs"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Rquisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //BUSQUEDA PERSONAL PROPUESTO
                                    selenium.Scroll("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']", PersonalPropuesto);
                                    Thread.Sleep(2000);
                                    //selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    Thread.Sleep(6000);
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Personal Porpuesto", true, file);
                                    Thread.Sleep(2000);
                                    //ADJUNTO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    Thread.Sleep(1000);
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(6000);
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
        public void SL_SolicitudRequisiciónMultiempresaTransferenciaContratoIndefinido()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_SolicitudRequisiciónMultiempresaTransferenciaContratoIndefinido")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string PersonalPropuesto = rows["PersonalPropuesto"].ToString();
                                string ddlCodEmprs = rows["ddlCodEmprs"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Rquisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //BUSQUEDA PERSONAL PROPUESTO
                                    selenium.Scroll("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']", PersonalPropuesto);
                                    Thread.Sleep(2000);
                                   // selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    Thread.Sleep(6000);
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleados_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Personal Porpuesto", true, file);
                                    Thread.Sleep(2000);
                                    //ADJUNTO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    Thread.Sleep(1000);
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(6000);
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
        public void SL_SolicitudRequisiciónMultiempresaConvocatoriaDirectaContratoFijo()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_SolicitudRequisiciónMultiempresaConvocatoriaDirectaContratoFijo")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string PersonalPropuesto = rows["PersonalPropuesto"].ToString();
                                string ddlCodEmprs = rows["ddlCodEmprs"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Rquisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //BUSQUEDA PERSONAL PROPUESTO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']", PersonalPropuesto);
                                    Thread.Sleep(2000);
                                    
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    Thread.Sleep(6000);
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleadosD_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleadosD_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Personal Porpuesto", true, file);
                                    Thread.Sleep(2000);
                                    //ADJUNTO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    Thread.Sleep(1000);
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //DETALLE
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(6000);
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
        public void SL_SolicitudRequisiciónMultiempresaConvocatoriaDirectaContratoIndefinido()
        {
            List<string> errorMessagesMetodo = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_SL.SL_SolicitudRequisiciónMultiempresaConvocatoriaDirectaContratoIndefinido")
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
                                //Datos Requisicion    
                                rows["GrupoRequi"].ToString().Length != 0 && rows["GrupoRequi"].ToString() != null &&
                                rows["ConvoRequi"].ToString().Length != 0 && rows["ConvoRequi"].ToString() != null &&
                                rows["FilSelRequi"].ToString().Length != 0 && rows["FilSelRequi"].ToString() != null &&
                                rows["CCostoRequi"].ToString().Length != 0 && rows["CCostoRequi"].ToString() != null &&
                                rows["PlazasRequi"].ToString().Length != 0 && rows["PlazasRequi"].ToString() != null &&
                                rows["CargoRequi"].ToString().Length != 0 && rows["CargoRequi"].ToString() != null &&
                                rows["MotivoRequi"].ToString().Length != 0 && rows["MotivoRequi"].ToString() != null &&
                                rows["ContratoRequi"].ToString().Length != 0 && rows["ContratoRequi"].ToString() != null &&
                                rows["TipoContratoRequi"].ToString().Length != 0 && rows["TipoContratoRequi"].ToString() != null &&
                                rows["SueldoRequi"].ToString().Length != 0 && rows["SueldoRequi"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ComentarioRequi"].ToString().Length != 0 && rows["ComentarioRequi"].ToString() != null &&
                                rows["PublicardadRequi"].ToString().Length != 0 && rows["PublicardadRequi"].ToString() != null &&
                                rows["MotivoCancelaRequi"].ToString().Length != 0 && rows["MotivoCancelaRequi"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                //Datos Requisicion                                               
                                string GrupoRequi = rows["GrupoRequi"].ToString();
                                string ConvoRequi = rows["ConvoRequi"].ToString();
                                string FilSelRequi = rows["FilSelRequi"].ToString();
                                string CCostoRequi = rows["CCostoRequi"].ToString();
                                string PlazasRequi = rows["PlazasRequi"].ToString();
                                string CargoRequi = rows["CargoRequi"].ToString();
                                string MotivoRequi = rows["MotivoRequi"].ToString();
                                string ContratoRequi = rows["ContratoRequi"].ToString();
                                string TipoContratoRequi = rows["TipoContratoRequi"].ToString();
                                string SueldoRequi = rows["SueldoRequi"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ComentarioRequi = rows["ComentarioRequi"].ToString();
                                string PublicardadRequi = rows["PublicardadRequi"].ToString();
                                string MotivoCancelaRequi = rows["MotivoCancelaRequi"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Funcionario = rows["Funcionario"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string PersonalPropuesto = rows["PersonalPropuesto"].ToString();
                                string ddlCodEmprs = rows["ddlCodEmprs"].ToString();
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
                                    string database = string.Empty;
                                    string User = string.Empty;
                                    string cod_empr = string.Empty;

                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        User = "ODESAR";
                                        cod_empr = "421";
                                    }
                                    else if (url.ToLower() == "http://dwtfsk:8093/".ToLower())

                                    {
                                        database = "SQL";
                                        User = "SDesar";
                                        cod_empr = "9";
                                    }
                                    //eliminar datos previos
                                    string eliminarSolicitud = $"Delete from NM_SOLTR where tip_apli ='R' AND COD_RESP ='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarSolicitud, database, User);

                                    string eliminarRequi = $"Delete from SL_REQPE where COD_EMPL ='{EmpleadoUser}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRequi, database, User);

                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //Process:Login//////////////////////////////////
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    ////////////////////////////////////////////////////

                                    //Process: Rquisision////////////////////////////////
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Requisición de Pe')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición de personal", true, file);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //NUEVO
                                    selenium.Click("//a[contains(@id,'btnNuevo')]");
                                    Thread.Sleep(2000);
                                    //GRUPO REQUISICION
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodGrse')]", GrupoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Requisición", true, file);
                                    //CONVOCATORIA
                                    selenium.Scroll("//select[contains(@id,'ddlForCobe')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlForCobe')]", ConvoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Convocatoria", true, file);
                                    //FILTRO
                                    if (ConvoRequi == "Convocatoria")
                                    {
                                        selenium.Scroll("//select[contains(@id,'ddlFilSele')]");
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFilSele')]", FilSelRequi);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Filtro", true, file);
                                    }
                                    Thread.Sleep(2000);
                                    //COSTO
                                    selenium.Scroll("//input[contains(@id,'txtCodCcos')]");
                                    selenium.Click("//input[contains(@id,'txtCodCcos')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'txtCodCcos')]", CCostoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Costo", true, file);
                                    Thread.Sleep(3000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//input[contains(@id,'txtNroPlaz')]");
                                    Thread.Sleep(3000);
                                    selenium.SendKeys("//input[contains(@id,'txtNroPlaz')]", PlazasRequi);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                    //CARGO
                                    selenium.Scroll("//select[contains(@id,'ddlCodCarp')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodCarp')]", CargoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    Thread.Sleep(2000);
                                    //MOTIVO
                                    selenium.Scroll("//select[contains(@id,'ddlCodMoti')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodMoti')]", MotivoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Motivo", true, file);
                                    Thread.Sleep(2000);
                                    //CONTRATO
                                    selenium.Scroll("//select[contains(@id,'ddlTipCont')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    Thread.Sleep(2000);
                                    //TIPO CONTRATO
                                    selenium.Scroll("//select[contains(@id, 'ddlContrDeta')]");
                                    selenium.SelectElementByName("//select[contains(@id, 'ddlContrDeta')]", TipoContratoRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo contrato", true, file);
                                    //CIUDAD
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //COMENTARIOS
                                    selenium.Scroll("//textarea[contains(@id,'KtxtDetRequ')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KtxtDetRequ')]", ComentarioRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comentario", true, file);
                                    Thread.Sleep(2000);
                                    //PUBLICAR SUELDO
                                    selenium.Scroll("//select[contains(@id,'ddlVisSuew')]");
                                    selenium.SelectElementByName("//select[contains(@id,'ddlVisSuew')]", PublicardadRequi);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Publicar sueldo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Requisición Personal", true, file);
                                    //FUNCIONARIO A REEMPLAZAR
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(6000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrlDivPoli_txtDivPoli')]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    selenium.SendKeys("//input[@type='search']", Funcionario);
                                    selenium.Screenshot("Funcionario a Reemplazar", true, file);
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    //BUSQUEDA PERSONAL PROPUESTO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtFiltroNombre']", PersonalPropuesto);
                                    Thread.Sleep(2000);
                                   
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnBuscar']");
                                    Thread.Sleep(6000);
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleadosD_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgFiltroEmpleadosD_ctl02_Lnkbfiltro']");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Personal Porpuesto", true, file);
                                    Thread.Sleep(2000);
                                    //ADJUNTO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtFiltroApellido']");
                                    Thread.Sleep(1000);
                                    for (int i = 0; i < 5; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(6000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    try
                                    {
                                        Thread.Sleep(6000);
                                        Screenshot("Alerta Registro Exitoso Requisición Personal", true, file);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Requisición", true, file);
                                    //CANCELAR REQUISICION
                                    Thread.Sleep(1000);
                                    selenium.Click("//button[contains(@id,'pLider')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'SELECCIÓN DE PERSONAL')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Requisición de Personal')]");
                                    Thread.Sleep(5000);
                                    //EMPRESA
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCodEmpr')]", ddlCodEmprs);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa Seleccionada", true, file);
                                    //DETALLE
                                    selenium.Click("//*[@id='printable']");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Click("//td[8]/a/i");
                                    Thread.Sleep(2500);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtMotCanc']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtMotCanc']", MotivoCancelaRequi);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Cancelar Req", true, file);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConCanc']");
                                    Thread.Sleep(2000);
                                    try
                                    {
                                        Screenshot("Alerta Cancelar Registro", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(8000);
                                        selenium.Screenshot("Requisición cancelada", true, file);

                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(6000);
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