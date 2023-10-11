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
    public class Self_Service_GCG : FuncionesVitales
    {

        string Modulo = "GCG";
        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();

        public Self_Service_GCG()
        {

        }
        [TestMethod]
        public void ED_AceptaciónReformulación()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SelfService.ED_AceptaciónReformulación")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["RMT_OMET"].ToString().Length != 0 && rows["RMT_OMET"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["ANO_OMET"].ToString().Length != 0 && rows["ANO_OMET"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string meta = rows["RMT_OMET"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string ano = rows["ANO_OMET"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "SQL")
                                    {
                                        string ActualizarReformulacion = $"update ED_REFOR set FEC_FINA = '31/12/2050' where RMT_OMET = '{meta}' and ANO_OMET = '{ano}'";
                                        db.UpdateDeleteInsert(ActualizarReformulacion, database, user);

                                        string ActualizarMeta = $"update ed_fomei set tip_meta = 'S' where COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(ActualizarMeta, database, user);

                                        string ActualizarTipoMeta = $"update ed_fomei set SEG_REFO = 'S' where COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(ActualizarTipoMeta, database, user);

                                        string Borrado = $"delete from ED_OBMEI where tip_meta = 'S' and cod_empr = 9 and ACT_USUA ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                    }

                                    //INGRESO A ENVALUACION DE OBJETIVOS ACEPTACION DE LA REFORMULACION
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Aceptación de la Reformulación')]");
                                    selenium.Click("//a[contains(.,'Aceptación de la Reformulación')]");
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Aceptación de la Reformulación", true, file);

                                    try
                                    {
                                        if (selenium.ExistControl("//*[@id=\'ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1\']"))
                                        {
                                            selenium.Screenshot("Detalle", true, file);

                                            //DETALLE REFORMULACION
                                            selenium.Click("//*[@id=\'ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1\']");
                                            Thread.Sleep(500);
                                            selenium.Screenshot("Detalles Aceptación de la Reformulación", true, file);

                                            Thread.Sleep(500);

                                            //OBSERVACIONES
                                            selenium.returnDriver().ExecuteScript("arguments[0].scrollIntoView(true);", selenium.returnDriver().FindElement(By.XPath("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']")));
                                            selenium.Click("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']");
                                            selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']", Observacion);

                                            //ACEPTAR REFORMULACION
                                            selenium.Click("//input[contains(@id,'btnAcepto')]");
                                            Thread.Sleep(500);
                                            Screenshot("Alerta Aceptación Reformulación", true, file);
                                            Thread.Sleep(3000);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(3000);
                                            Screenshot("Reformulación Aceptada con Éxito", true, file);
                                            selenium.AcceptAlert();

                                        }
                                        else
                                        {
                                            selenium.Screenshot("Sin datos de reformulacion", true, file);
                                        }

                                    }
                                    catch (Exception e)
                                    {

                                        continue;
                                    }

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
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
        public void ED_AceptaciónFormulaciónMetas()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SelfService.ED_AceptaciónFormulaciónMetas")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["RMT_OMET"].ToString().Length != 0 && rows["RMT_OMET"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["ANO_OMET"].ToString().Length != 0 && rows["ANO_OMET"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string meta = rows["RMT_OMET"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string ano = rows["ANO_OMET"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "SQL")
                                    {
                                        string ActualizarReformulacion = $"update ED_REFOR set FEC_FINA = '31/12/2050' where RMT_OMET = '{meta}' and ANO_OMET = '{ano}'";
                                        db.UpdateDeleteInsert(ActualizarReformulacion, database, user);

                                        string ActualizarMeta = $"update ed_fomei set tip_meta = 'S' where COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(ActualizarMeta, database, user);

                                        string ActualizarTipoMeta = $"update ed_fomei set SEG_REFO = 'N' where COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(ActualizarTipoMeta, database, user);

                                        string Borrado = $"delete from ED_OBMEI where tip_meta = 'S' and cod_empr = 9 and ACT_USUA ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                    }

                                    //INGRESO A EVALUACION DE OBJETIVOS//ACEPTACION DE LA FORMULACION METAS
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'EF Aceptacion Formulacion de Metas')]");
                                    selenium.Click("//a[contains(.,'EF Aceptacion Formulacion de Metas')]");
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Aceptación de la Formulación Metas", true, file);

                                    try
                                    {
                                        if (selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i"))
                                        {
                                            selenium.Screenshot("Detalle", true, file);

                                            //DETALLE FORMULACION METAS
                                            selenium.Click("//*[@id=\'ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1\']");
                                            Thread.Sleep(1500);
                                            selenium.Screenshot("Detalles Aceptación de la Formulación", true, file);
                                            Thread.Sleep(1500);

                                            //OBSERVACIONES
                                            selenium.returnDriver().ExecuteScript("arguments[0].scrollIntoView(true);", selenium.returnDriver().FindElement(By.XPath("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']")));
                                            selenium.Click("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']");
                                            selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']", Observacion);

                                            //ACEPTAR FORMULACION
                                            selenium.Click("//input[contains(@id,'btnAcepto')]");
                                            Thread.Sleep(1500);
                                            Screenshot("Alerta Aceptación Reformulación", true, file);
                                            Thread.Sleep(3000);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(3000);
                                            Screenshot("Reformulación Aceptada con Éxito", true, file);
                                            selenium.AcceptAlert();

                                        }
                                        else
                                        {
                                            selenium.Screenshot("Sin datos de reformulacion", true, file);
                                        }

                                    }
                                    catch (Exception e)
                                    {

                                        continue;
                                    }

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
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
        public void ED_AprobacionReformulaciónMetas()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SelfService.ED_AprobacionReformulaciónMetas")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["RMT_OMET"].ToString().Length != 0 && rows["RMT_OMET"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["ANO_OMET"].ToString().Length != 0 && rows["ANO_OMET"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string meta = rows["RMT_OMET"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string ano = rows["ANO_OMET"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "SQL")
                                    {
                                        string ActualizarReformulacion = $"update ED_REFOR set FEC_FINA = '31/12/2050' where RMT_OMET = '{meta}' and ANO_OMET = '{ano}'";
                                        db.UpdateDeleteInsert(ActualizarReformulacion, database, user);

                                        string ActualizarMeta = $"update ed_fomei set tip_meta = 'A' where COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(ActualizarMeta, database, user);

                                        string ActualizarTipoMeta = $"update ed_fomei set SEG_REFO = 'S' where COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(ActualizarTipoMeta, database, user);

                                        string Borrado = $"delete from ED_OBMEI where tip_meta = 'A' and cod_empr = 9 and ACT_USUA ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                    }

                                    //INGRESO A ENVALUACION DE OBJETIVOS ACEPTACION DE LA REFORMULACION
                                    //ROL LIDER
                                    selenium.Click("//*[@id='pLider']");
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(1500);
                                    selenium.Scroll("//a[contains(.,'EF Aprobacion Reformulacion de Metas')]");
                                    selenium.Click("//a[contains(.,'EF Aprobacion Reformulacion de Metas')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Aprobación de la Reformulación", true, file);

                                    try
                                    {
                                        if (selenium.ExistControl("//input[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']"))
                                        {
                                            selenium.Screenshot("Detalle", true, file);

                                            //DETALLE REFORMULACION
                                            selenium.Click("//input[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']");
                                            Thread.Sleep(2500);
                                            selenium.Screenshot("Detalles Aceptación de la Reformulación", true, file);
                                            Thread.Sleep(2500);

                                            //OBSERVACIONES
                                            selenium.ScrollTo("0","800");
                                            Thread.Sleep(2500);
                                            selenium.Click("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']");
                                            selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']", Observacion);

                                            //APROBAR REFORMULACION
                                            selenium.Click("//input[@id='ctl00_ContenidoPagina_btnAproMet']");
                                            Thread.Sleep(500);
                                            Screenshot("Alerta Aprobación Reformulación", true, file);
                                            Thread.Sleep(3000);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(3000);
                                            Screenshot("Reformulación Aceptada con Éxito", true, file);
                                            selenium.AcceptAlert();

                                        }
                                        else
                                        {
                                            selenium.Screenshot("Sin datos de reformulacion", true, file);
                                        }

                                    }
                                    catch (Exception e)
                                    {

                                        continue;
                                    }

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
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
        public void ED_AprobacionFormulaciónMetas()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SelfService.ED_AprobacionFormulaciónMetas")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["RMT_OMET"].ToString().Length != 0 && rows["RMT_OMET"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["ANO_OMET"].ToString().Length != 0 && rows["ANO_OMET"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string meta = rows["RMT_OMET"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string ano = rows["ANO_OMET"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "SQL")
                                    {
                                        string ActualizarReformulacion = $"update ED_REFOR set FEC_FINA = '31/12/2050' where RMT_OMET = '{meta}' and ANO_OMET = '{ano}'";
                                        db.UpdateDeleteInsert(ActualizarReformulacion, database, user);

                                        string ActualizarMeta = $"update ed_fomei set tip_meta = 'A' where COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(ActualizarMeta, database, user);

                                        string ActualizarTipoMeta = $"update ed_fomei set SEG_REFO = 'N' where COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(ActualizarTipoMeta, database, user);

                                        string Borrado = $"delete from ED_OBMEI where tip_meta = 'N' and cod_empr = 9 and ACT_USUA ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                    }

                                    //INGRESO A ENVALUACION DE OBJETIVOS ACEPTACION DE LA REFORMULACION
                                    //ROL LIDER
                                    selenium.Click("//*[@id='pLider']");
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(1500);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[5]/ul/li[10]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[5]/ul/li[10]/a");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Aprobación de la Reformulación", true, file);

                                    try
                                    {
                                        if (selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i"))
                                        {
                                            selenium.Screenshot("Detalle", true, file);

                                            //DETALLE REFORMULACION
                                            selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i");
                                            Thread.Sleep(2500);
                                            selenium.Screenshot("Detalles Aceptación de la Formulación", true, file);
                                            Thread.Sleep(2500);

                                            //OBSERVACIONES
                                            selenium.ScrollTo("0", "800");
                                            Thread.Sleep(2500);
                                            selenium.Click("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']");
                                            selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']", Observacion);

                                            //APROBAR REFORMULACION
                                            selenium.Click("//input[@id='ctl00_ContenidoPagina_btnAproMet']");
                                            Thread.Sleep(500);
                                            Screenshot("Alerta Aprobación Formulacion", true, file);
                                            Thread.Sleep(3000);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(3000);
                                            Screenshot("Formulacion Aceptada con Éxito", true, file);
                                            selenium.AcceptAlert();

                                        }
                                        else
                                        {
                                            selenium.Screenshot("Sin datos de reformulacion", true, file);
                                        }

                                    }
                                    catch (Exception e)
                                    {

                                        continue;
                                    }

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
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
        public void ED_FormulaciónObjetivosMetas()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_FormulaciónObjetivosMetas")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Perspectiva"].ToString().Length != 0 && rows["Perspectiva"].ToString() != null &&
                                rows["Estrategico"].ToString().Length != 0 && rows["Estrategico"].ToString() != null &&
                                rows["Areas"].ToString().Length != 0 && rows["Areas"].ToString() != null &&
                                rows["Indicador"].ToString().Length != 0 && rows["Indicador"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Perspectiva = rows["Perspectiva"].ToString();
                                string Estrategico = rows["Estrategico"].ToString();
                                string Areas = rows["Areas"].ToString();
                                string Indicador = rows["Indicador"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //Parametrización de Formulación Metas para el usuario 
                                    if (database == "SQL")
                                    {
                                        string actualizar = $"update ED_FOMEI set SEG_REFO = 'N', TIP_META = 'F' where COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(actualizar, database, user);
                                        Thread.Sleep(2000);
                                        string actualizar2 = $"update ED_FORMU set FEC_FINA = '31/12/2050' WHERE RMT_OMET = '10002'";
                                        db.UpdateDeleteInsert(actualizar2, database, user);
                                        Thread.Sleep(2000);
                                    }
                                 
                                    //INGRESO A FORMULACION DE OBJETIVOS
                                    selenium.Scroll("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Formulación Objetivos')]");
                                    selenium.Click("//a[contains(.,'Formulación Objetivos')]");
                                    selenium.Screenshot("Formulación de objetivos", true, file);
                                    Thread.Sleep(1000);
                                    //DETALLE
                                    selenium.Click("//*[@id=\"ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1\"]");
                                    Thread.Sleep(1000);
                                    //BORRAR FORMULACIONES
                                    if (selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_gvLista_ctl03_btnRemover']/i"))
                                    {
                                        for (int i = 0; i < 3; i++)
                                        {
                                            selenium.Scroll("//a[@id='ctl00_ContenidoPagina_gvLista_ctl03_btnRemover']/i");
                                            selenium.Click("//a[@id='ctl00_ContenidoPagina_gvLista_ctl03_btnRemover']/i");
                                            Thread.Sleep(3000);
                                            selenium.AcceptAlert();
                                        }
                                    }
                                    //INGRESO A FORMULACION DE OBJETIVOS
                                    selenium.Scroll("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Formulación Objetivos')]");
                                    selenium.Click("//a[contains(.,'Formulación Objetivos')]");
                                    selenium.Screenshot("Formulación de objetivos", true, file);
                                    Thread.Sleep(1000);
                                    if (selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i"))
                                    {
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i");
                                        Thread.Sleep(1000);

                                        for (int i = 0; i < 3; i++)
                                        {
                                            //PERSPECTIVA
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_COD_PERS']", Perspectiva);
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Perspectiva ", true, file);
                                            //ESTRATEGICO
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_COD_OBES']", Estrategico);
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Estrategico ", true, file);
                                            //AREAS
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_COD_OBAR']", Areas);
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Áreas ", true, file);
                                            selenium.ScrollTo("0", "500");
                                            //META INDIVIDUAL
                                            selenium.Click("//textarea[@id='ctl00_ContenidoPagina_KtxtDesMeta_txtTexto']");
                                            Thread.Sleep(2000);
                                            selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KtxtDesMeta_txtTexto']", "Metas Individuales");
                                            Thread.Sleep(2000);
                                            //INDICADOR
                                            if (database == "SQL")
                                            {
                                                selenium.ScrollTo("0", "800");
                                                selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_COD_INDI']", Indicador);
                                                Thread.Sleep(2000);
                                                selenium.Screenshot("Indicador ", true, file);
                                            }

                                            //SEGUIMIENTO DATOS
                                            Thread.Sleep(2000);
                                            selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_PRO_PTRI']", "10");
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Datos Seguimiento ", true, file);

                                            if (i == 0)
                                            {
                                                selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_POR_PESO']", "50");
                                                Thread.Sleep(2000);
                                                selenium.Screenshot("Peso ", true, file);
                                                selenium.Click("//input[@id='ctl00_ContenidoPagina_btnAplicar']");
                                                Thread.Sleep(2000);
                                                selenium.Scroll("//a[@id='ctl00_ContenidoPagina_gvLista_ctl03_btnRemover']/i");
                                                Thread.Sleep(1000);
                                                selenium.Screenshot("Aplicado ", true, file);
                                                Thread.Sleep(1000);
                                                selenium.ScrollTo("0", "200");
                                            }
                                            else
                                            {
                                                selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_POR_PESO']", "25");
                                                Thread.Sleep(2000);
                                                selenium.Screenshot("Peso ", true, file);
                                                selenium.Click("//input[@id='ctl00_ContenidoPagina_btnAplicar']");
                                                Thread.Sleep(2000);
                                                selenium.Scroll("//a[@id='ctl00_ContenidoPagina_gvLista_ctl03_btnRemover']/i");
                                                Thread.Sleep(1000);
                                                selenium.Screenshot("Aplicado ", true, file);
                                                Thread.Sleep(1000);
                                                selenium.ScrollTo("0", "200");
                                            }
                                        }

                                        //OBSERVACIONES
                                        selenium.Scroll("//a[@id='ctl00_ContenidoPagina_gvLista_ctl03_btnRemover']/i");
                                        Thread.Sleep(1000);
                                        selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']", "PRUEBACALIDAD");
                                        Thread.Sleep(3000);
                                        selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnEnvJefe']");
                                        //ENVIAR
                                        selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_btnEnvJefe')]");
                                        Thread.Sleep(3000);
                                        Screenshot("Registro Correcto", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(3000);
                                        Screenshot("Registro Exitoso", true, file);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(2000);
                                        selenium.Close();
                                        fv.ConvertWordToPDF(file, database);
                                        Thread.Sleep(2000);
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
        public void ED_InformeComparaciónCompetenciasVsMetasObjetivos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_InformeComparaciónCompetenciasVsMetasObjetivos")
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
                                //Datos Necesidades Formación    
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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    File.Delete(@"C:\Users\kactusscm\Downloads\ReporteGenerado1.pdf");
                                    File.Delete(@"C:\Users\kactusscm\Downloads\ReporteGenerado2.pdf");

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A INFORME 
                                    //ROL LIDER
                                    selenium.Click("//*[@id='pLider']");
                                    Thread.Sleep(1500);
                                    selenium.Scroll("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(1500);
                                    selenium.Scroll("//a[contains(.,'Informe de comparación por Competencias Vs Metas y Objetivos - Mapa de Talentos')]");
                                    selenium.Click("//a[contains(.,'Informe de comparación por Competencias Vs Metas y Objetivos - Mapa de Talentos')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Informe de comparación por Competencias Vs Metas y Objetivos", true, file);
                                    Thread.Sleep(1000);
                                    //DETALLE
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Detalle Informe", true, file);
                                    //REPORTE SIN CHECK
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_DtgFomeiR_ctl02_LinkButton1']/i");
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_DtgFomeiR_ctl02_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    //INFOME GENERADO
                                    selenium.Screenshot("Informe Generado", true, file);
                                    Thread.Sleep(2000);
                                    //GENERAR PDF
                                    selenium.Click("//a[@id='ctl00_btnImprimirPDFToolbar']");
                                    Thread.Sleep(20000);
                                    Screenshot("Imprimir PDF", true, file);
                                    //GUARDAR PDF
                                    for (int i = 0; i < 5; i++)
                                    {
                                        Keyboard.SendKeys("{TAB}");
                                        Thread.Sleep(1000);
                                    }

                                    Keyboard.SendKeys("{ENTER}");
                                    Thread.Sleep(5000);
                                    Keyboard.SendKeys("{DOWN}");
                                    Thread.Sleep(5000);
                                    Keyboard.SendKeys("{ENTER}");
                                    Thread.Sleep(5000);

                                    for (int i = 0; i < 5; i++)
                                    {
                                        Keyboard.SendKeys("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    Keyboard.SendKeys("{ENTER}");
                                    Thread.Sleep(5000);
                                    Keyboard.SendKeys("ReporteGenerado1");
                                    Thread.Sleep(5000);
                                    Screenshot("Guardar", true, file);
                                    Keyboard.SendKeys("{ENTER}");
                                    Thread.Sleep(5000);

                                    //REPORTE INDIVIDUAL
                                    selenium.Click("//a[@id='ctl00_btnRetornar']");
                                    Thread.Sleep(6000);
                                    //DETALLE
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Detalle Informe", true, file);
                                    //REPORTE CHECK
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_DtgFomeiR_ctl02_ChkRegiFomei']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_DtgFomeiR_ctl02_ChkRegiFomei']");
                                    Thread.Sleep(2000);
                                    //REPORTE
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_DtgFomeiR_ctl02_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    //INFOME GENERADO
                                    selenium.Screenshot("Informe Generado", true, file);
                                    Thread.Sleep(2000);
                                    //GENERAR PDF
                                    selenium.Click("//a[@id='ctl00_btnImprimirPDFToolbar']");
                                    Thread.Sleep(20000);
                                    Screenshot("PDF", true, file);
                                    //GUARDAR
                                    Keyboard.SendKeys("{ENTER}");
                                    Thread.Sleep(5000);
                                    Keyboard.SendKeys("ReporteGenerado2");
                                    Thread.Sleep(5000);
                                    Screenshot("Guardar", true, file);
                                    Keyboard.SendKeys("{ENTER}");
                                    Thread.Sleep(6000);

                                    //ABRIR PDF DESCARGADO 1
                                    string pdfPath = Path.Combine(Application.StartupPath, @"C:\Users\kactusscm\Downloads\ReporteGenerado1.pdf");
                                    Process.Start(pdfPath);
                                    Thread.Sleep(10000);
                                    Screenshot("PDF 1 ABIERTO", true, file);
                                    //ABRIR PDF DESCARGADO 2
                                    string pdfPath1 = Path.Combine(Application.StartupPath, @"C:\Users\kactusscm\Downloads\ReporteGenerado2.pdf");
                                    Process.Start(pdfPath1);
                                    Thread.Sleep(10000);
                                    Screenshot("PDF 2 ABIERTO", true, file);
                                    LimpiarProcesos();
                                    Thread.Sleep(6000);
                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);
                                    
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
        public void ED_InformeMetasObjetivos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_InformeMetasObjetivos")
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
                                //Datos Necesidades Formación    
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
                                string ano = rows["ano"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    File.Delete(@"C:\Users\kactusscm\Downloads\Reportedesempenometas.pdf");
                                    //PARAMETRIZACION
                                    string actualizar = $"update ED_FOMEI set SEG_REFO = 'N', TIP_META = 'T' where COD_EMPL = '{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(actualizar, database, user);
                                    Thread.Sleep(2000);
                                    //INICIO
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    //INGRESO A INFORME 
                                    selenium.Scroll("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(1500);
                                    selenium.Scroll("//*[@id='MenuContex']/div[2]/div[1]/ul/li[8]/ul/li[16]/a");
                                    selenium.Click("//*[@id='MenuContex']/div[2]/div[1]/ul/li[8]/ul/li[16]/a");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Informe de Metas y Objetivos", true, file);
                                    Thread.Sleep(1000);
                                    //CHECK REPORTE DESEMPEÑO METAS
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_chkIndRep1']");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Reporte Desempeño Metas", true, file);
                                    //AÑO
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAnoMeta']", ano);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Año", true, file);
                                    //CONSULTAR
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnConsul']");
                                    //CHECK EMPLEADO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_dtgBiEmple_ctl02_chkSelEmpl']");
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_dtgBiEmple_ctl02_chkSelEmpl']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empleado Consultado", true, file);
                                    //GENERAR
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGenerar']");
                                    Thread.Sleep(10000);
                                    selenium.Screenshot("Informe Generado", true, file);
                                    //IMPRIMIR
                                    selenium.Click("//a[@id='ctl00_btnImprimir']");
                                    Thread.Sleep(10000);
                                    Screenshot("Imprimir PDF", true, file);
                                    Keyboard.SendKeys("{ESC}");
                                    Thread.Sleep(5000);
                                    //PDF
                                    selenium.Click("//a[@id='ctl00_btnImprimirPDFToolbar']");
                                    Thread.Sleep(10000);
                                    Screenshot("Descargar PDF", true, file);
                                    //ABRIR PDF DESCARGADO 
                                    string pdfPath = Path.Combine(Application.StartupPath, @"C:\Users\kactusscm\Downloads\Reportedesempenometas.pdf");
                                    Process.Start(pdfPath);
                                    Thread.Sleep(10000);
                                    Screenshot("PDF ABIERTO", true, file);
                                    LimpiarProcesos();
                                    Thread.Sleep(6000);
                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);

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
        public void ED_ReformulaciónMetas()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_ReformulaciónMetas")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["ruta"].ToString().Length != 0 && rows["ruta"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["TipoAdjunto"].ToString().Length != 0 && rows["TipoAdjunto"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string ruta = rows["ruta"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string TipoAdjunto = rows["TipoAdjunto"].ToString();
                                string url = rows["url"].ToString();


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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "SQL")
                                    {
                                        string ActualizarMeta = $"update ed_fomei set tip_meta = 'F' where COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(ActualizarMeta, database, user);

                                        string ActualizarTipoMeta = $"update ed_fomei set SEG_REFO = 'S' where COD_EMPL = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(ActualizarTipoMeta, database, user);

                                        string Borrado = $"delete from ED_OBMEI where tip_meta = 'F' and cod_empr = 9 and ACT_USUA ='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                    }

                                    //INGRESO A EVALUACION DE OBJETIVOS//REFOMRULACION METAS
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(1500);
                                    selenium.Scroll("//a[contains(@href, 'frmEdFomeiRFL.ASPX')]");
                                    selenium.Click("//a[contains(@href, 'frmEdFomeiRFL.ASPX')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Reformulación Metas y Objetivos", true, file);

                                    try
                                    {
                                        if (selenium.ExistControl("//input[@id='ctl00_ContenidoPagina_gvLista2_ctl03_btnVer']"))
                                        {
                                            selenium.Screenshot("Detalle", true, file);

                                            //DETALLE FORMULACION METAS
                                            selenium.Click("//input[@id='ctl00_ContenidoPagina_gvLista2_ctl03_btnVer']");
                                            Thread.Sleep(1500);
                                            selenium.Screenshot("Detalles Reformulación", true, file);
                                            Thread.Sleep(1500);
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_KCtrTipoDocumento_ddlTIP_DOCU']", TipoAdjunto);
                                            Thread.Sleep(5000);
                                            //OBSERVACIONES
                                            selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnAdjuntar']");
                                            Thread.Sleep(3000);
                                            selenium.Click("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']");
                                            Thread.Sleep(3000);
                                            selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']", Observacion);
                                            Thread.Sleep(3000);

                                            //ENVIAR JEFE
                                            selenium.Click("//input[@id='ctl00_ContenidoPagina_btnEnvJefe']");
                                            Thread.Sleep(1500);
                                            Screenshot("Alerta Reformulación", true, file);
                                            Thread.Sleep(3000);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(3000);
                                            Screenshot("Reformulación Registrada con Éxito", true, file);
                                            selenium.AcceptAlert();

                                        }
                                        else
                                        {
                                            selenium.Screenshot("Sin datos de reformulacion", true, file);
                                        }

                                    }
                                    catch (Exception e)
                                    {

                                        continue;
                                    }

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
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
        public void ED_ReporteMetasIndividualesxEmpleado()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_ReporteMetasIndividualesxEmpleado")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["ano"].ToString().Length != 0 && rows["ano"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string ano = rows["ano"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    File.Delete(@"C: \Users\kactusscm\Downloads\ReporteMetasGenerado.pdf");

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A REPORTE METAS INDIVIDUALES POR EMPLEADO
                                    //ROL LIDER
                                    selenium.Click("//*[@id='pLider']");
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(1500);
                                    selenium.Scroll("//a[contains(.,'Reporte Metas Individuales x Empleados')]");
                                    selenium.Click("//a[contains(.,'Reporte Metas Individuales x Empleados')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Reporte Metas Individuales x Empleados", true, file);
                                    //AÑO A CONSULTAR
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_txtAño']");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtAño']", ano);
                                    Thread.Sleep(1500);
                                    //CONSULTAR
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_btnConsulYea']/span");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Registro Consultado", true, file);

                                    try
                                    {
                                        if (selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl02_LinkButton1']/i"))
                                        {
                                            selenium.Screenshot("Detalle", true, file);

                                            //DETALLE INFORME
                                            selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl02_LinkButton1']/i");
                                            Thread.Sleep(2500);
                                            selenium.Screenshot("Informe Generado", true, file);
                                            Thread.Sleep(2500);
                                            //IMPRIMIR
                                            Thread.Sleep(2500);
                                            selenium.Click("//a[@id='ctl00_btnImprimir']");
                                            Thread.Sleep(30000);
                                            Screenshot("Imprimir PDF", true, file);
                                            //GUARDAR PDF
                                            for (int i = 0; i < 5; i++)
                                            {
                                                Keyboard.SendKeys("{TAB}");
                                                Thread.Sleep(1000);
                                            }

                                            Keyboard.SendKeys("{ENTER}");
                                            Thread.Sleep(5000);
                                            Keyboard.SendKeys("{DOWN}");
                                            Thread.Sleep(5000);
                                            Keyboard.SendKeys("{ENTER}");
                                            Thread.Sleep(5000);

                                            for (int i = 0; i < 5; i++)
                                            {
                                                Keyboard.SendKeys("{TAB}");
                                                Thread.Sleep(1000);
                                            }
                                            Keyboard.SendKeys("{ENTER}");
                                            Thread.Sleep(5000);
                                            Keyboard.SendKeys("ReporteMetasGenerado");
                                            Thread.Sleep(5000);
                                            Screenshot("Guardar", true, file);
                                            Keyboard.SendKeys("{ENTER}");
                                            Thread.Sleep(5000);
                                            
                                            //ABRIR PDF DESCARGADO 1
                                            string pdfPath = Path.Combine(Application.StartupPath, @"C:\Users\kactusscm\Downloads\ReporteMetasGenerado.pdf");
                                            Process.Start(pdfPath);
                                            Thread.Sleep(10000);
                                            Screenshot("PDF ABIERTO", true, file);
                                            Thread.Sleep(5000);
                                            LimpiarProcesos();


                                        }
                                        else
                                        {
                                            selenium.Screenshot("Sin datos de reporte", true, file);
                                        }

                                    }
                                    catch (Exception e)
                                    {

                                        continue;
                                    }

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
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
        public void ED_ReporteMetasIndividuales()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_ReporteMetasIndividuales")
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
                                //Datos Necesidades Formación    
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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    File.Delete(@"C:\Users\kactusscm\Downloads\ReporteMetasIndividualesGenerado.pdf");
                                    
                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A REPORTE METAS INDIVIDUALES
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(1500);
                                    selenium.Scroll("(//a[contains(@href, 'frmEdFomeiRL.aspx')])[2]");
                                    selenium.Click("(//a[contains(@href, 'frmEdFomeiRL.aspx')])[2]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Reporte Metas Individuales", true, file);
                                    try
                                    {
                                        if (selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i"))
                                        {
                                            selenium.Screenshot("Detalle", true, file);

                                            //DETALLE INFORME
                                            selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i");
                                            Thread.Sleep(2500);
                                            selenium.Screenshot("Reporte Metas Individuales Generado", true, file);
                                            Thread.Sleep(2500);
                                            //IMPRIMIR
                                            Thread.Sleep(2500);
                                            selenium.Click("//a[@id='ctl00_btnImprimir']");
                                            Thread.Sleep(30000);
                                            Screenshot("Imprimir PDF", true, file);
                                            //GUARDAR PDF
                                            for (int i = 0; i < 5; i++)
                                            {
                                                Keyboard.SendKeys("{TAB}");
                                                Thread.Sleep(1000);
                                            }

                                            Keyboard.SendKeys("{ENTER}");
                                            Thread.Sleep(5000);
                                            Keyboard.SendKeys("{DOWN}");
                                            Thread.Sleep(5000);
                                            Keyboard.SendKeys("{ENTER}");
                                            Thread.Sleep(5000);

                                            for (int i = 0; i < 5; i++)
                                            {
                                                Keyboard.SendKeys("{TAB}");
                                                Thread.Sleep(1000);
                                            }
                                            Keyboard.SendKeys("{ENTER}");
                                            Thread.Sleep(5000);
                                            Keyboard.SendKeys("ReporteMetasIndividualesGenerado");
                                            Thread.Sleep(5000);
                                            Screenshot("Guardar", true, file);
                                            Keyboard.SendKeys("{ENTER}");
                                            Thread.Sleep(5000);

                                            //ABRIR PDF DESCARGADO 1
                                            string pdfPath = Path.Combine(Application.StartupPath, @"C:\Users\kactusscm\Downloads\ReporteMetasIndividualesGenerado.pdf");
                                            Process.Start(pdfPath);
                                            Thread.Sleep(10000);
                                            Screenshot("PDF ABIERTO", true, file);
                                            Thread.Sleep(5000);
                                            LimpiarProcesos();


                                        }
                                        else
                                        {
                                            selenium.Screenshot("Sin datos de reporte", true, file);
                                        }

                                    }
                                    catch (Exception e)
                                    {

                                        continue;
                                    }

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
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
        public void ED_RatificacionFormulacionMetas()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_RatificacionFormulacionMetas")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string Empleado= rows["Empleado"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "SQL")
                                    {
                                        string ActualizarMeta = $"update ed_fomei set tip_meta = 'R' where COD_EMPL = '{Empleado}'";
                                        db.UpdateDeleteInsert(ActualizarMeta, database, user);

                                        string ActualizarTipoMeta = $"update ed_fomei set SEG_REFO = 'N' where COD_EMPL = '{Empleado}'";
                                        db.UpdateDeleteInsert(ActualizarTipoMeta, database, user);
                                    }

                                    //INGRESO A Ratificacion Formulacion Metas
                                    //ROL LIDER
                                    selenium.Click("//*[@id='pLider']");
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(1500);
                                    selenium.Scroll("(//a[contains(@href, 'frmLIEdFomeiTL.aspx')])[2]");
                                    selenium.Click("(//a[contains(@href, 'frmLIEdFomeiTL.aspx')])[2]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ratificacion Formulacion Metas", true, file);

                                    try
                                    {
                                        if (selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i"))
                                        {
                                            selenium.Screenshot("Detalle", true, file);

                                            //DETALLE REFORMULACION
                                            selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i");
                                            Thread.Sleep(2500);
                                            selenium.Screenshot("Detalles Ratificacion Formulacion Metas", true, file);
                                            Thread.Sleep(2500);

                                            //OBSERVACIONES
                                            selenium.ScrollTo("0", "800");
                                            Thread.Sleep(2500);
                                            selenium.Click("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']");
                                            selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']", Observacion);

                                            //APROBAR METAS
                                            selenium.Click("//input[@id='ctl00_ContenidoPagina_btnAproMet']");
                                            Thread.Sleep(500);
                                            Screenshot("Alerta Aceptacion Metas", true, file);
                                            Thread.Sleep(3000);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(3000);
                                            Screenshot("Ratificacion Formulacion Metas con Éxito", true, file);
                                            selenium.AcceptAlert();

                                        }
                                        else
                                        {
                                            selenium.Screenshot("Sin datos de Ratificacion Formulacion Metas", true, file);
                                        }

                                    }
                                    catch (Exception e)
                                    {

                                        continue;
                                    }

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
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
        public void ED_RatificacionEvaluacionMetas()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_RatificacionEvaluacionMetas")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string Empleado = rows["Empleado"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string url = rows["url"].ToString();
                                string dato1 = rows["dato1"].ToString();
                                string dato2 = rows["dato2"].ToString();
                                string dato3 = rows["dato3"].ToString();


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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "SQL")
                                    {
                                        string ActualizarMeta = $"update ed_fomei set tip_meta = 'V' where COD_EMPL = '{Empleado}'";
                                        db.UpdateDeleteInsert(ActualizarMeta, database, user);

                                        string ActualizarTipoMeta = $"update ed_fomei set SEG_REFO = 'N' where COD_EMPL = '{Empleado}'";
                                        db.UpdateDeleteInsert(ActualizarTipoMeta, database, user);
                                    }

                                    //INGRESO A Ratificacion Formulacion Metas
                                    //ROL LIDER
                                    selenium.Click("//*[@id='pLider']");
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(1500);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[5]/ul/li[13]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[5]/ul/li[13]/a");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ratificacion Evaluacion Metas", true, file);

                                    try
                                    {
                                        if (selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i"))
                                        {
                                            selenium.Screenshot("Detalle", true, file);

                                            //DETALLE REFORMULACION
                                            selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']/i");
                                            Thread.Sleep(2500);
                                            if (selenium.ExistControl("//a[@id='ctl00_btnCerrar']/i"))
                                            {
                                                selenium.Click("//a[@id='ctl00_btnCerrar']/i");
                                                Thread.Sleep(2000);
                                            }
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Detalle Ratificacion Evaluacion Metas", true, file);
                                            Thread.Sleep(2500);

                                            //CALIFICAR METAS
                                            selenium.ScrollTo("0", "800");
                                            Thread.Sleep(2500);
                                            //DATO 1
                                            selenium.Click("//input[@id='ctl00_ContenidoPagina_gvLista_ctl02_EJE_CUTA']");
                                            Thread.Sleep(2500);
                                            selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_gvLista_ctl02_EJE_CUTA']", dato1);
                                            //DATO 2
                                            selenium.Click("//input[@id='ctl00_ContenidoPagina_gvLista_ctl03_EJE_CUTA']");
                                            Thread.Sleep(2500);
                                            selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_gvLista_ctl03_EJE_CUTA']", dato2);
                                            //DATO 2
                                            selenium.Click("//input[@id='ctl00_ContenidoPagina_gvLista_ctl04_EJE_CUTA']");
                                            Thread.Sleep(2500);
                                            selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_gvLista_ctl04_EJE_CUTA']", dato3);
                                            Thread.Sleep(2500);
                                            selenium.Screenshot("Evaluacion Ingresada", true, file);
                                            //OBSERVACIONES
                                            selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']",Observacion);
                                            Thread.Sleep(2500);
                                            selenium.Screenshot("Observaciones", true, file);
                                            //GUARDAR RATIFICACION
                                            selenium.Click("//input[@id='ctl00_ContenidoPagina_btnSolCali']");
                                            Thread.Sleep(500);
                                            Screenshot("Alerta Ratificacion Metas", true, file);
                                            Thread.Sleep(3000);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(3000);
                                            Screenshot("Ratificacion Evaluacion Metas con Éxito", true, file);
                                            selenium.AcceptAlert();

                                        }
                                        else
                                        {
                                            selenium.Screenshot("Sin datos de Ratificacion Evaluacion Metas", true, file);
                                        }

                                    }
                                    catch (Exception e)
                                    {

                                        continue;
                                    }

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
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
        public void ED_RatificacionReformulacionMetas()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_RatificacionReformulacionMetas")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string Empleado = rows["Empleado"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "SQL")
                                    {
                                        string ActualizarMeta = $"update ed_fomei set tip_meta = 'R' where COD_EMPL = '{Empleado}'";
                                        db.UpdateDeleteInsert(ActualizarMeta, database, user);

                                        string ActualizarTipoMeta = $"update ed_fomei set SEG_REFO = 'S' where COD_EMPL = '{Empleado}'";
                                        db.UpdateDeleteInsert(ActualizarTipoMeta, database, user);
                                    }

                                    //INGRESO A Ratificacion ReFormulacion Metas
                                    //ROL LIDER
                                    selenium.Click("//*[@id='pLider']");
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE OBJETIVOS')]");
                                    Thread.Sleep(1500);
                                    selenium.Scroll("//div[@id='MenuContex']/div[2]/div/ul/li[5]/ul/li[15]/a");
                                    selenium.Click("//div[@id='MenuContex']/div[2]/div/ul/li[5]/ul/li[15]/a");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ratificacion Reformulacion Metas", true, file);

                                    try
                                    {
                                        if (selenium.ExistControl("//input[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']"))
                                        {
                                            selenium.Screenshot("Detalle", true, file);

                                            //DETALLE REFORMULACION
                                            selenium.Click("//input[@id='ctl00_ContenidoPagina_dtgDatos_ctl03_LinkButton1']");
                                            Thread.Sleep(2500);
                                            selenium.Screenshot("Detalles Ratificacion Reformulacion Metas", true, file);
                                            Thread.Sleep(2500);

                                            //OBSERVACIONES
                                            selenium.ScrollTo("0", "800");
                                            Thread.Sleep(2500);
                                            selenium.Click("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']");
                                            selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_OBS_ERVA']", Observacion);

                                            //APROBAR REFORMULACION METAS
                                            selenium.Click("//input[@id='ctl00_ContenidoPagina_btnAproMet']");
                                            Thread.Sleep(500);
                                            Screenshot("Alerta Reformulacion Metas", true, file);
                                            Thread.Sleep(3000);
                                            selenium.AcceptAlert();
                                            Thread.Sleep(3000);
                                            Screenshot("Ratificacion Reformulacion Metas con Éxito", true, file);
                                            selenium.AcceptAlert();

                                        }
                                        else
                                        {
                                            selenium.Screenshot("Sin datos de Ratificacion Reformulacion Metas", true, file);
                                        }

                                    }
                                    catch (Exception e)
                                    {

                                        continue;
                                    }

                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);

                                    Thread.Sleep(2000);
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
        public void ED_GCG360ValidarGuardarCompromiso()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_GCG360ValidarGuardarCompromiso")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Aspecto"].ToString().Length != 0 && rows["Aspecto"].ToString() != null &&
                                rows["Comportamiento"].ToString().Length != 0 && rows["Comportamiento"].ToString() != null &&
                                rows["Compromiso"].ToString().Length != 0 && rows["Compromiso"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["Recurso"].ToString().Length != 0 && rows["Recurso"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Aspecto = rows["Aspecto"].ToString();
                                string Comportamiento = rows["Comportamiento"].ToString();
                                string Compromiso = rows["Compromiso"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string Recurso = rows["Recurso"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    if (database == "SQL")
                                    {
                                        string Borrado1 = $"delete from ED_PERAP where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado1, database, user);
                                        string Borrado = $"delete from ED_COMPR where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                        string Borrado2 = $"delete from GD_TIACC where ACT_USUA = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado2, database, user);
                                    }
                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //ROL LIDER
                                    selenium.Click("//span[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    //DEFINICION PLANES DESARROLLO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    selenium.Screenshot("Definicion Planes Desarrollo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    Thread.Sleep(2000);
                                    //LISTADO COMPROMISOS PENDIENTES
                                    selenium.Scroll("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Listado Compromisos pendientes", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle Registro Compromisos", true, file);
                                    Thread.Sleep(2000);
                                    //REGISTRO COMPROMISOS
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAspEval']", Aspecto);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Aspecto", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDetAspEval']", Comportamiento);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comportamiento", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDesComp']", Compromiso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtOtrCompr_txtTexto']", "OTRO COMPROMISO");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso Editado", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtRecRequ_txtTexto']", Recurso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Recurso", true, file);
                                    Thread.Sleep(2000);
                                    //GUARDAR COMPROMISO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Guardar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Compromiso Guardado", true, file);
                                    //CERRAR CONCERTACION
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnCerCompro']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cerrar Concertacion", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnCerCompro']");
                                    Thread.Sleep(4000);
                                    Screenshot("Alerta Cerrado Concertacion", true, file);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Proceso Terminado Correctamente", true, file);
                                    Thread.Sleep(4000);
                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
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
        public void ED_GCG360ValidarCamposObligatoriosGuardarCompromiso()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_GCG360ValidarCamposObligatoriosGuardarCompromiso")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Aspecto"].ToString().Length != 0 && rows["Aspecto"].ToString() != null &&
                                rows["Comportamiento"].ToString().Length != 0 && rows["Comportamiento"].ToString() != null &&
                                rows["Compromiso"].ToString().Length != 0 && rows["Compromiso"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["Recurso"].ToString().Length != 0 && rows["Recurso"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Aspecto = rows["Aspecto"].ToString();
                                string Comportamiento = rows["Comportamiento"].ToString();
                                string Compromiso = rows["Compromiso"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string Recurso = rows["Recurso"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    if (database == "SQL")
                                    {
                                        string Borrado1 = $"delete from ED_PERAP where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado1, database, user);
                                        string Borrado = $"delete from ED_COMPR where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                        string Borrado2 = $"delete from GD_TIACC where ACT_USUA = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado2, database, user);
                                    }
                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //ROL LIDER
                                    selenium.Click("//span[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    //DEFINICION PLANES DESARROLLO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    selenium.Screenshot("Definicion Planes Desarrollo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    Thread.Sleep(2000);
                                    //LISTADO COMPROMISOS PENDIENTES
                                    selenium.Scroll("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Listado Compromisos pendientes", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle Registro Compromisos", true, file);
                                    Thread.Sleep(2000);
                                    //REGISTRO COMPROMISOS
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    //IR A BOTON GUARDAR SIN LLENAR DATOS, VALIDAR CAMPOS OBLIGARORIOS
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Guardar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(4000);
                                    selenium.ScrollTo("0", "200");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Validacion Campos Obligatorios 1", true, file);
                                    Thread.Sleep(2000);
                                    //LLENAR DATOS MENOS RECURSOS Y GUARDAR, VALIDAR CAMPOS OBLIGATORIOS
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAspEval']", Aspecto);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Aspecto", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDetAspEval']", Comportamiento);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comportamiento", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDesComp']", Compromiso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtOtrCompr_txtTexto']", "OTRO COMPROMISO");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso Editado", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    //GUARDAR COMPROMISO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Guardar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(4000);
                                    selenium.ScrollTo("0", "600");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Validacion Campos Obligatorios 2", true, file);
                                    //LLENA RESTANTE CAMPOS Y SE GUARDA COMPROMISO CORRECTAMENTE
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtRecRequ_txtTexto']", Recurso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Recurso", true, file);
                                    Thread.Sleep(2000);
                                    //GUARDAR COMPROMISO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Guardar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Compromiso Guardado", true, file);
                                    //CERRAR CONCERTACION
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnCerCompro']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cerrar Concertacion", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnCerCompro']");
                                    Thread.Sleep(4000);
                                    Screenshot("Alerta Cerrado Concertacion", true, file);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Proceso Terminado Correctamente", true, file);
                                    Thread.Sleep(4000);
                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
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
        public void ED_GCG360ValidarPersonaApoyoGuardarCompromiso()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_GCG360ValidarPersonaApoyoGuardarCompromiso")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Aspecto"].ToString().Length != 0 && rows["Aspecto"].ToString() != null &&
                                rows["Comportamiento"].ToString().Length != 0 && rows["Comportamiento"].ToString() != null &&
                                rows["Compromiso"].ToString().Length != 0 && rows["Compromiso"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["Recurso"].ToString().Length != 0 && rows["Recurso"].ToString() != null &&
                                rows["PersonaApoyo"].ToString().Length != 0 && rows["PersonaApoyo"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Aspecto = rows["Aspecto"].ToString();
                                string Comportamiento = rows["Comportamiento"].ToString();
                                string Compromiso = rows["Compromiso"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string Recurso = rows["Recurso"].ToString();
                                string url = rows["url"].ToString();
                                string PersonaApoyo = rows["PersonaApoyo"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    if (database == "SQL")
                                    {
                                        string Borrado1 = $"delete from ED_PERAP where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado1, database, user);
                                        string Borrado = $"delete from ED_COMPR where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                        string Borrado2 = $"delete from GD_TIACC where ACT_USUA = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado2, database, user);
                                    }
                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //ROL LIDER
                                    selenium.Click("//span[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    //DEFINICION PLANES DESARROLLO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    selenium.Screenshot("Definicion Planes Desarrollo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    Thread.Sleep(2000);
                                    //LISTADO COMPROMISOS PENDIENTES
                                    selenium.Scroll("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Listado Compromisos pendientes", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle Registro Compromisos", true, file);
                                    Thread.Sleep(2000);
                                    //REGISTRO COMPROMISOS
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAspEval']", Aspecto);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Aspecto", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDetAspEval']", Comportamiento);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comportamiento", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDesComp']", Compromiso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtOtrCompr_txtTexto']", "OTRO COMPROMISO");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso Editado", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtRecRequ_txtTexto']", Recurso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Recurso", true, file);
                                    Thread.Sleep(2000);
                                    //PERSONA APOYO AGREGAR
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtConsuCedEmpl']", PersonaApoyo);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Persona Apoyo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_btnConsulCed']/span");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Persona Consultada", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgEmpleadosApoy_ctl03_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Persona Apoyo Agregada", true, file);
                                    Thread.Sleep(2000);
                                    //GUARDAR COMPROMISO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Guardar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Compromiso Guardado", true, file);
                                    //CERRAR CONCERTACION
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnCerCompro']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cerrar Concertacion", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnCerCompro']");
                                    Thread.Sleep(4000);
                                    Screenshot("Alerta Cerrado Concertacion", true, file);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Proceso Terminado Correctamente", true, file);
                                    Thread.Sleep(4000);
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    //-------------------------------------------------PARTE 2---AGREGA Y ELIMINA LA PERSONA APOYO ANTES DE GUARDAR COMPROMISO---------------------------------------------------------
                                    if (database == "SQL")
                                    {
                                        string Borrado1 = $"delete from ED_PERAP where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado1, database, user);
                                        string Borrado = $"delete from ED_COMPR where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                    }
                                    Thread.Sleep(10000);

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    //ROL LIDER
                                    selenium.Click("//span[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    //DEFINICION PLANES DESARROLLO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    selenium.Screenshot("Definicion Planes Desarrollo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    Thread.Sleep(2000);
                                    //LISTADO COMPROMISOS PENDIENTES
                                    selenium.Scroll("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Listado Compromisos pendientes", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle Registro Compromisos", true, file);
                                    Thread.Sleep(2000);
                                    //REGISTRO COMPROMISOS
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAspEval']", Aspecto);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Aspecto", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDetAspEval']", Comportamiento);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comportamiento", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDesComp']", Compromiso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtRecRequ_txtTexto']", Recurso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Recurso", true, file);
                                    Thread.Sleep(2000);
                                    //PERSONA APOYO AGREGAR
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtConsuCedEmpl']", PersonaApoyo);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Persona Apoyo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_btnConsulCed']/span");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Persona Consultada", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgEmpleadosApoy_ctl03_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Persona Apoyo Agregada", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgPerApoy_ctl03_btnElimi']/i");
                                    Thread.Sleep(2000);
                                    Screenshot("Eliminar Persona Apoyo antes Guardar compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Persona Eliminada", true, file);
                                    Thread.Sleep(2000);
                                    //GUARDAR COMPROMISO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Guardar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Compromiso Guardado", true, file);
                                    //CERRAR CONCERTACION
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnCerCompro']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cerrar Concertacion", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnCerCompro']");
                                    Thread.Sleep(4000);
                                    Screenshot("Alerta Cerrado Concertacion", true, file);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Proceso Terminado Correctamente", true, file);
                                    Thread.Sleep(4000);
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);
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
        public void ED_GCG360ValidarEliminarPersonaApoyoGuardarCompromiso()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_GCG360ValidarEliminarPersonaApoyoGuardarCompromiso")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Aspecto"].ToString().Length != 0 && rows["Aspecto"].ToString() != null &&
                                rows["Comportamiento"].ToString().Length != 0 && rows["Comportamiento"].ToString() != null &&
                                rows["Compromiso"].ToString().Length != 0 && rows["Compromiso"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["Recurso"].ToString().Length != 0 && rows["Recurso"].ToString() != null &&
                                rows["PersonaApoyo"].ToString().Length != 0 && rows["PersonaApoyo"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Aspecto = rows["Aspecto"].ToString();
                                string Comportamiento = rows["Comportamiento"].ToString();
                                string Compromiso = rows["Compromiso"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string Recurso = rows["Recurso"].ToString();
                                string url = rows["url"].ToString();
                                string PersonaApoyo = rows["PersonaApoyo"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    if (database == "SQL")
                                    {
                                        string Borrado1 = $"delete from ED_PERAP where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado1, database, user);
                                        string Borrado = $"delete from ED_COMPR where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                        string Borrado2 = $"delete from GD_TIACC where ACT_USUA = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado2, database, user);
                                    }
                                    Thread.Sleep(2000);

                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    //ROL LIDER
                                    selenium.Click("//span[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    //DEFINICION PLANES DESARROLLO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    selenium.Screenshot("Definicion Planes Desarrollo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    Thread.Sleep(2000);
                                    //LISTADO COMPROMISOS PENDIENTES
                                    selenium.Scroll("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Listado Compromisos pendientes", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle Registro Compromisos", true, file);
                                    Thread.Sleep(2000);
                                    //REGISTRO COMPROMISOS
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAspEval']", Aspecto);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Aspecto", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDetAspEval']", Comportamiento);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comportamiento", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDesComp']", Compromiso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtOtrCompr_txtTexto']", "OTRO COMPROMISO");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso Editado", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtRecRequ_txtTexto']", Recurso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Recurso", true, file);
                                    Thread.Sleep(2000);
                                    //PERSONA APOYO AGREGAR
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtConsuCedEmpl']", PersonaApoyo);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Persona Apoyo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_btnConsulCed']/span");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Persona Consultada", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgEmpleadosApoy_ctl03_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Persona Apoyo Agregada", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgPerApoy_ctl03_btnElimi']/i");
                                    Thread.Sleep(2000);
                                    Screenshot("Eliminar Persona Apoyo antes Guardar compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Persona Eliminada", true, file);
                                    Thread.Sleep(2000);
                                    //GUARDAR COMPROMISO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Guardar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Compromiso Guardado", true, file);
                                    //CERRAR CONCERTACION
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnCerCompro']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cerrar Concertacion", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnCerCompro']");
                                    Thread.Sleep(4000);
                                    Screenshot("Alerta Cerrado Concertacion", true, file);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Proceso Terminado Correctamente", true, file);
                                    Thread.Sleep(4000);
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);
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
        public void ED_GCG360ValidarEdicionCompromiso()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_GCG360ValidarEdicionCompromiso")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Aspecto"].ToString().Length != 0 && rows["Aspecto"].ToString() != null &&
                                rows["Comportamiento"].ToString().Length != 0 && rows["Comportamiento"].ToString() != null &&
                                rows["Compromiso"].ToString().Length != 0 && rows["Compromiso"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["Recurso"].ToString().Length != 0 && rows["Recurso"].ToString() != null &&
                                rows["Aspecto2"].ToString().Length != 0 && rows["Aspecto2"].ToString() != null &&
                                rows["Comportamiento2"].ToString().Length != 0 && rows["Comportamiento2"].ToString() != null &&
                                rows["Compromiso2"].ToString().Length != 0 && rows["Compromiso2"].ToString() != null &&
                                rows["Recurso2"].ToString().Length != 0 && rows["Recurso2"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Aspecto = rows["Aspecto"].ToString();
                                string Comportamiento = rows["Comportamiento"].ToString();
                                string Compromiso = rows["Compromiso"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string Recurso = rows["Recurso"].ToString();
                                string Aspecto2 = rows["Aspecto2"].ToString();
                                string Comportamiento2 = rows["Comportamiento2"].ToString();
                                string Compromiso2 = rows["Compromiso2"].ToString();
                                string Recurso2 = rows["Recurso2"].ToString();
                                string CampoCompromiso2 =rows["CampoCompromiso2"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    if (database == "SQL")
                                    {
                                        string Borrado1 = $"delete from ED_PERAP where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado1, database, user);
                                        string Borrado = $"delete from ED_COMPR where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                        string Borrado2 = $"delete from GD_TIACC where ACT_USUA = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado2, database, user);
                                    }
                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //ROL LIDER
                                    selenium.Click("//span[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    //DEFINICION PLANES DESARROLLO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    selenium.Screenshot("Definicion Planes Desarrollo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    Thread.Sleep(2000);
                                    //LISTADO COMPROMISOS PENDIENTES
                                    selenium.Scroll("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Listado Compromisos pendientes", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle Registro Compromisos", true, file);
                                    Thread.Sleep(2000);
                                    //REGISTRO COMPROMISOS
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAspEval']", Aspecto);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Aspecto", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDetAspEval']", Comportamiento);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comportamiento", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDesComp']", Compromiso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtOtrCompr_txtTexto']", "OTRO COMPROMISO");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso Editado", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtRecRequ_txtTexto']", Recurso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Recurso", true, file);
                                    Thread.Sleep(2000);
                                    //GUARDAR COMPROMISO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Guardar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Compromiso Guardado", true, file);
                                    //EDITAR COMPROMISO
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina2_DgListadoEspe_ctl03_btnEdit']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso a editar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina2_DgListadoEspe_ctl03_btnEdit']/i");
                                    //EDITAR DATOS REGISTRO
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAspEval']", Aspecto2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Aspecto", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDetAspEval']", Comportamiento2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comportamiento", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDesComp']", Compromiso2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtOtrCompr_txtTexto']", CampoCompromiso2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso Editado", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtRecRequ_txtTexto']", Recurso2);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Recurso", true, file);
                                    Thread.Sleep(2000);
                                    //ACTUALIZAR COMPROMISO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnActualizar']");
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnActualizar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Actualizado", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina2_DgListadoEspe_ctl03_btnEdit']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso editado", true, file);
                                    Thread.Sleep(4000);
                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
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
        public void ED_GCG360ValidarEliminarCompromiso()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_GCG360ValidarEliminarCompromiso")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Aspecto"].ToString().Length != 0 && rows["Aspecto"].ToString() != null &&
                                rows["Comportamiento"].ToString().Length != 0 && rows["Comportamiento"].ToString() != null &&
                                rows["Compromiso"].ToString().Length != 0 && rows["Compromiso"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["Recurso"].ToString().Length != 0 && rows["Recurso"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Aspecto = rows["Aspecto"].ToString();
                                string Comportamiento = rows["Comportamiento"].ToString();
                                string Compromiso = rows["Compromiso"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string Recurso = rows["Recurso"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    if (database == "SQL")
                                    {
                                        string Borrado1 = $"delete from ED_PERAP where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado1, database, user);
                                        string Borrado = $"delete from ED_COMPR where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                        string Borrado2 = $"delete from GD_TIACC where ACT_USUA = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado2, database, user);
                                    }
                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //ROL LIDER
                                    selenium.Click("//span[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    //DEFINICION PLANES DESARROLLO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    selenium.Screenshot("Definicion Planes Desarrollo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    Thread.Sleep(2000);
                                    //LISTADO COMPROMISOS PENDIENTES
                                    selenium.Scroll("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Listado Compromisos pendientes", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle Registro Compromisos", true, file);
                                    Thread.Sleep(2000);
                                    //REGISTRO COMPROMISOS
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAspEval']", Aspecto);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Aspecto", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDetAspEval']", Comportamiento);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comportamiento", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDesComp']", Compromiso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtOtrCompr_txtTexto']", "OTRO COMPROMISO");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso Editado", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtRecRequ_txtTexto']", Recurso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Recurso", true, file);
                                    Thread.Sleep(2000);
                                    //GUARDAR COMPROMISO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Guardar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Compromiso Guardado", true, file);
                                    //ELIMINAR COMPROMISO
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina2_DgListadoEspe_ctl03_btnElimi']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso a editar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina2_DgListadoEspe_ctl03_btnElimi']/i");
                                    Thread.Sleep(2000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso Eliminado", true, file);
                                    Thread.Sleep(4000);
                                    fv.ConvertWordToPDF(file, database);
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
        public void ED_GCG360ValidarCompromisoTipoOtroMayor99GuardarCompromiso()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_GCG360ValidarCompromisoTipoOtroMayor99GuardarCompromiso")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Aspecto"].ToString().Length != 0 && rows["Aspecto"].ToString() != null &&
                                rows["Comportamiento"].ToString().Length != 0 && rows["Comportamiento"].ToString() != null &&
                                rows["Compromiso"].ToString().Length != 0 && rows["Compromiso"].ToString() != null &&
                                rows["CampoCompromiso"].ToString().Length != 0 && rows["CampoCompromiso"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["Recurso"].ToString().Length != 0 && rows["Recurso"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Aspecto = rows["Aspecto"].ToString();
                                string Comportamiento = rows["Comportamiento"].ToString();
                                string Compromiso = rows["Compromiso"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string Recurso = rows["Recurso"].ToString();
                                string url = rows["url"].ToString();
                                string CampoCompromiso = rows["CampoCompromiso"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    if (database == "SQL")
                                    {
                                        string Borrado1 = $"delete from ED_PERAP where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado1, database, user);
                                        string Borrado = $"delete from ED_COMPR where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                        string Borrado2 = $"delete from GD_TIACC where ACT_USUA = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado2, database, user);
                                    }
                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //ROL LIDER
                                    selenium.Click("//span[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    //DEFINICION PLANES DESARROLLO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    selenium.Screenshot("Definicion Planes Desarrollo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    Thread.Sleep(2000);
                                    //LISTADO COMPROMISOS PENDIENTES
                                    selenium.Scroll("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Listado Compromisos pendientes", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle Registro Compromisos", true, file);
                                    Thread.Sleep(2000);
                                    //REGISTRO COMPROMISOS
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAspEval']", Aspecto);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Aspecto", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDetAspEval']", Comportamiento);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comportamiento", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDesComp']", Compromiso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtOtrCompr_txtTexto']", CampoCompromiso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso Editado", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtRecRequ_txtTexto']", Recurso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Recurso", true, file);
                                    Thread.Sleep(2000);
                                    //GUARDAR COMPROMISO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Guardar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Compromiso Guardado", true, file);
                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
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
        public void ED_GCG360ValidarGuardarCompromisoDigiflagN()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_GCG360ValidarGuardarCompromisoDigiflagN")
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
                                //Datos Necesidades Formación    
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Aspecto"].ToString().Length != 0 && rows["Aspecto"].ToString() != null &&
                                rows["Comportamiento"].ToString().Length != 0 && rows["Comportamiento"].ToString() != null &&
                                rows["Compromiso"].ToString().Length != 0 && rows["Compromiso"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["Recurso"].ToString().Length != 0 && rows["Recurso"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null

                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Aspecto = rows["Aspecto"].ToString();
                                string Comportamiento = rows["Comportamiento"].ToString();
                                string Compromiso = rows["Compromiso"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string Recurso = rows["Recurso"].ToString();
                                string url = rows["url"].ToString();

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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    if (database == "SQL")
                                    {
                                        string digi = $"update GN_DIGIF set VAL_VARI='N' where COD_VARI='K0000063'";
                                        db.UpdateDeleteInsert(digi, database, user);
                                        string Borrado1 = $"delete from ED_PERAP where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado1, database, user);
                                        string Borrado = $"delete from ED_COMPR where rmt_licc = 3172";
                                        db.UpdateDeleteInsert(Borrado, database, user);
                                        string Borrado2 = $"delete from GD_TIACC where ACT_USUA = '{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(Borrado2, database, user);
                                    }
                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //ROL LIDER
                                    selenium.Click("//span[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    //DEFINICION PLANES DESARROLLO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    selenium.Screenshot("Definicion Planes Desarrollo", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrSolicitud_ctl09_ImageButton2']/i");
                                    Thread.Sleep(2000);
                                    //LISTADO COMPROMISOS PENDIENTES
                                    selenium.Scroll("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Listado Compromisos pendientes", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle Registro Compromisos", true, file);
                                    Thread.Sleep(2000);
                                    //REGISTRO COMPROMISOS
                                    selenium.Click("//button[contains(.,'Cerrar')]");
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlAspEval']", Aspecto);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Aspecto", true, file);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDetAspEval']", Comportamiento);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comportamiento", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KCtrlTxtDesComp_txtTexto']", "COMPROMISO DIGIFLAG N");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecInic_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_KCtrlFecFina_txtFecha']", Fecha);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fechas Compromiso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtRecRequ_txtTexto']", Recurso);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Recurso", true, file);
                                    Thread.Sleep(2000);
                                    //GUARDAR COMPROMISO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Guardar", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_btnGuardar']");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Compromiso Guardado", true, file);
                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    if (database == "SQL")
                                    {
                                        string digi = $"update GN_DIGIF set VAL_VARI='S' where COD_VARI='K0000063'";
                                        db.UpdateDeleteInsert(digi, database, user);
                                       
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
        public void ED_GCG360ValidarLiderNoObserveSuDDefinicionCompromisos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.Self_Service_GCG.ED_GCG360ValidarLiderNoObserveSuDDefinicionCompromisos")
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
                                //Datos Necesidades Formación    
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
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    if (database == "SQL")
                                    {
                                        string digi = $"update GN_DIGIF set VAL_VARI='N' where COD_VARI='K0000063'";
                                        db.UpdateDeleteInsert(digi, database, user);

                                    }
                                    //INICIO PROGRAMA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //ROL LIDER
                                    selenium.Click("//span[contains(.,'Rol Lider')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Rol Lider", true, file);
                                    //DEFINICION PLANES DESARROLLO INDIVIDUAL
                                    selenium.Scroll("//a[contains(.,'EVALUACIÓN DE COMPETENCIAS')]");
                                    selenium.Click("//a[contains(.,'EVALUACIÓN DE COMPETENCIAS')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'EF Definición Planes de Desarrollo')]");
                                    selenium.Click("//a[contains(.,'EF Definición Planes de Desarrollo')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Definicion Planes Desarrollo", true, file);
                                    Thread.Sleep(2000);
                                    //DETALLES, NO SE DEBE VISUALIZAR EL DEL USUARIO 507195
                                    selenium.Scroll("//table[@id='tblListado']/tbody/tr/td[8]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Validacion Lider No Observa su Definicion", true, file);
                                    Thread.Sleep(2000);
                                    fv.ConvertWordToPDF(file, database);
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    if (database == "SQL")
                                    {
                                        string digi = $"update GN_DIGIF set VAL_VARI='S' where COD_VARI='K0000063'";
                                        db.UpdateDeleteInsert(digi, database, user);

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


