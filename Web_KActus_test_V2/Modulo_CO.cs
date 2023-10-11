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
    public class Modulo_CO : FuncionesVitales
    {

        string Modulo = "Modulo_CO";
        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();

        public Modulo_CO()
        {

        }
        [TestMethod]
        public void CO_VisualizaciónEncuestas()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_CO.CO_VisualizaciónEncuestas")
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
                                    if (database == "SQL")
                                    {
                                        string EliminarRegistro = $"delete from CO_ENCUE where cod_empr = 9 and cod_vari = 9";
                                        db.UpdateDeleteInsert(EliminarRegistro, database, user);
                                    }
                                    else
                                    {
                                        string EliminarRegistro = $"delete from CO_ENCUE where cod_empr = 421 and cod_vari = 8";
                                        db.UpdateDeleteInsert(EliminarRegistro, database, user);

                                    }

                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A ENCUESTAS
                                    selenium.Click("//a[contains(.,'ENCUESTAS')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Encuestas')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A ENCUESTAS", true, file);

                                    //MODULO
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_ModuloCO']/div/div[2]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Clima Organizacional", true, file);
                                    if (database == "SQL")
                                    {
                                        //ENCUESTA
                                        selenium.Click("//div[@id='NDJ8TnxDT3w0fDIyICAgICAgICAgIA==']/div");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Encuesta", true, file);
                                    }
                                    else
                                    {
                                        //ENCUESTA
                                        selenium.Click("//div[@id='NXxOfENPfDQwMDQ1MDIzfDE=']/span[2]");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Encuesta", true, file);
                                    }

                                    //DETALLE
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgVaria_ctl02_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Visualizacion Encuesta", true, file);

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
        public void CO_DiligenciarEncuestasInserciónResultados()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_CO.CO_DiligenciarEncuestasInserciónResultados")
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
                                rows["Respuesta1"].ToString().Length != 0 && rows["Respuesta1"].ToString() != null &&
                                rows["Respuesta2"].ToString().Length != 0 && rows["Respuesta2"].ToString() != null &&
                                rows["Respuesta3"].ToString().Length != 0 && rows["Respuesta3"].ToString() != null &&
                                rows["Respuesta4"].ToString().Length != 0 && rows["Respuesta4"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string url = rows["url"].ToString();
                                string Respuesta1 = rows["Respuesta1"].ToString();
                                string Respuesta2 = rows["Respuesta2"].ToString();
                                string Respuesta3 = rows["Respuesta3"].ToString();
                                string Respuesta4 = rows["Respuesta4"].ToString();
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
                                        string EliminarRegistro = $"delete from CO_ENCUE where cod_empr = 9 and cod_vari = 9";
                                        db.UpdateDeleteInsert(EliminarRegistro, database, user);
                                    }
                                    else
                                    {
                                        string EliminarRegistro = $"delete from CO_ENCUE where cod_empr = 421 and cod_vari = 8";
                                        db.UpdateDeleteInsert(EliminarRegistro, database, user);

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

                                    //INGRESO A ENCUESTAS
                                    selenium.Click("//a[contains(.,'ENCUESTAS')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Encuestas')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A ENCUESTAS", true, file);
                                    //MODULO
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_ModuloCO']/div/div[2]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Clima Organizacional", true, file);
                                    if (database == "SQL")
                                    {
                                        //ENCUESTA
                                        selenium.Click("//div[@id='NDJ8TnxDT3w0fDIyICAgICAgICAgIA==']/div");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Encuesta", true, file);
                                    }
                                    else
                                    {
                                        //ENCUESTA
                                        selenium.Click("//div[@id='NXxOfENPfDQwMDQ1MDIzfDE=']/span[2]");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Encuesta", true, file);
                                    }
                                    
                                    //DETALLE
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgVaria_ctl02_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Visualizacion Encuesta", true, file);
                                    //PREGUNTA 1
                                    selenium.Click("//label[contains(.,'Totalmente de acuerdo')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//button[@id='btnSiguiente']");
                                    selenium.Screenshot("Pregunta 1", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//button[@id='btnSiguiente']");
                                    Thread.Sleep(2000);
                                    //PREGUNTA 2
                                    selenium.Click("//label[contains(.,'Totalmente de acuerdo')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//button[@id='btnSiguiente']");
                                    selenium.Screenshot("Pregunta 2", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//button[@id='btnSiguiente']");
                                    Thread.Sleep(2000);
                                    //PREGUNTA 3
                                    selenium.Click("//label[contains(.,'Parcialmente de acuerdo')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//button[@id='btnSiguiente']");
                                    selenium.Screenshot("Pregunta 3", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//button[@id='btnSiguiente']");
                                    Thread.Sleep(2000);
                                    //PREGUNTA 4
                                    selenium.Click("//label[contains(.,'Parcialmente de acuerdo')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//button[@id='btnSiguiente']");
                                    selenium.Screenshot("Pregunta 4", true, file);
                                    Thread.Sleep(2000);
                                    //GUARDAR
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_divPreguntas']/div/div[3]/div/div[2]/button");
                                    Thread.Sleep(2000);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Encuesta Diligenciada", true, file);
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
        public void CO_VisualizaciónResultados()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_CO.CO_VisualizaciónResultados")
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
                                    string username = Environment.UserName;
                                    //CREACION DOCUMENTO
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);

                                    //PARAMETRIZACION PREVIA


                                    //-------------------------------------------------------INICIO PRUEBA---------------------------------------------------------------------------------                               

                                    //LOGIN
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    //INGRESO A RESULTADOS
                                    selenium.Click("//a[contains(.,'ENCUESTAS')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Resultado de Clima Organizacional')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("INGRESO A RESULTADOS", true, file);

                                    //SELECCIONAR LISTADO
                                    selenium.Click("//td[4]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("LISTADO", true, file);

                                    //RESULTADO
                                    selenium.Screenshot("RESULTADOS", true, file);
                                    Thread.Sleep(2000);


                                    //REPORTE TOTAL
                                    selenium.Click("//a[contains(.,'ENCUESTAS')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Resultado de Clima Organizacional')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//td[5]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("REPORTE TOTAL", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Total')]");
                                    Thread.Sleep(4000);
                                    //REPORTE TOTAL
                                    
                                    String mainWin = selenium.MainWindow();
                                    String modalWin = selenium.PopupWindow();
                                    if (selenium.CountWindow() == 2)
                                    {
                                        selenium.ChangeWindow(modalWin);
                                        Thread.Sleep(5000);
                                        selenium.MaximizeWindow();
                                        Thread.Sleep(500);
                                        selenium.Screenshot("REPORTE TOTAL", true, file);
                                        //IMPRIMIR
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_imprimir']");
                                        Thread.Sleep(6000);
                                        Screenshot("IMPRIMIR REPORTE TOTAL", true, file);
                                        Thread.Sleep(4000);
                                        SendKeys.SendWait("{ESC}");
                                        Thread.Sleep(2000);
                                        //PDF
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_imprimirPDF']");
                                        Thread.Sleep(6000);
                                        Screenshot("PDF REPORTE TOTAL", true, file);

                                        //ABRIR PDF
                                        string pdfPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\Reporte_Valoracion_competencias.pdf");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(6000);
                                        Screenshot("PDF ABIERTO", true, file);
                                        Thread.Sleep(5000);
                                        KillProcesos("Acrobat.exe");
                                        Thread.Sleep(10000);
                                        File.Delete(@"C:\Users\" + username + @"\Downloads\Reporte_Valoracion_competencias.pdf");
                                        selenium.Close();
                                    }
                                    selenium.ChangeWindow(mainWin);
                                    //REPORTE CENTRO COSTOS
                                    selenium.Click("//a[contains(.,'ENCUESTAS')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Resultado de Clima Organizacional')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//td[5]/a/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("REPORTE CENTRO COSTOS", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(text(),'Centro de Costo')]");
                                    Thread.Sleep(4000);
                                    //REPORTE TOTAL
                                   
                                    String mainWin1 = selenium.MainWindow();
                                    String modalWin1 = selenium.PopupWindow();
                                    if (selenium.CountWindow() == 2)
                                    {
                                        selenium.ChangeWindow(modalWin1);
                                        Thread.Sleep(5000);
                                        selenium.MaximizeWindow();
                                        Thread.Sleep(500);
                                        selenium.Screenshot("REPORTE CENTRO COSTOS", true, file);
                                        //IMPRIMIR
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_imprimir']");
                                        Thread.Sleep(6000);
                                        Screenshot("IMPRIMIR REPORTE TOTAL", true, file);
                                        Thread.Sleep(4000);
                                        SendKeys.SendWait("{ESC}");
                                        Thread.Sleep(2000);
                                        //PDF
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_imprimirPDF']");
                                        Thread.Sleep(6000);
                                        Screenshot("PDF REPORTE TOTAL", true, file);

                                        //ABRIR PDF
                                        string pdfPath = Path.Combine(Application.StartupPath, @"C:\Users\" + username + @"\Downloads\Reporte_Valoracion_competencias.pdf");
                                        Process.Start(pdfPath);
                                        Thread.Sleep(6000);
                                        Screenshot("PDF ABIERTO", true, file);
                                        Thread.Sleep(5000);
                                        KillProcesos("Acrobat.exe");
                                        Thread.Sleep(10000);
                                        File.Delete(@"C:\Users\" + username + @"\Downloads\Reporte_Valoracion_competencias.pdf");
                                        selenium.Close();
                                    }
                                    selenium.ChangeWindow(mainWin1);

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