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
using APITest;
using System.IO;
using System.Management;
using System.Diagnostics;
using System.Data;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;

namespace Web_Kactus_Test
{
    /// <summary>
    /// Descripción resumida de PruebaPrueba
    /// </summary>
    [CodedUITest]
    public class SmartPeople_NTC_6 : FuncionesVitales
    {
        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();
        public SmartPeople_NTC_6()
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
        public void SmartPeople_frmAcropueNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_6.SmartPeople_frmAcropueNTC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                //rows["ActualizarEsp"].ToString().Length != 0 && rows["ActualizarEsp"].ToString() != null &&

                                rows["NumSolicitudEsp"].ToString().Length != 0 && rows["NumSolicitudEsp"].ToString() != null &&
                                rows["EstadoEsp"].ToString().Length != 0 && rows["EstadoEsp"].ToString() != null &&
                                rows["EmpresaEsp"].ToString().Length != 0 && rows["EmpresaEsp"].ToString() != null &&

                                rows["DatosColaboradorEsp"].ToString().Length != 0 && rows["DatosColaboradorEsp"].ToString() != null &&
                                rows["IdentificacionColabEsp"].ToString().Length != 0 && rows["IdentificacionColabEsp"].ToString() != null &&
                                rows["NombresColabEsp"].ToString().Length != 0 && rows["NombresColabEsp"].ToString() != null &&
                                rows["ApellidosColabEsp"].ToString().Length != 0 && rows["ApellidosColabEsp"].ToString() != null &&
                                rows["ResponsableSolicitudEsp"].ToString().Length != 0 && rows["ResponsableSolicitudEsp"].ToString() != null &&
                                rows["IdentificacionRespEsp"].ToString().Length != 0 && rows["IdentificacionRespEsp"].ToString() != null &&
                                rows["NombreRespEsp"].ToString().Length != 0 && rows["NombreRespEsp"].ToString() != null &&
                                rows["ApellidosRespEsp"].ToString().Length != 0 && rows["ApellidosRespEsp"].ToString() != null &&
                                rows["DatosSolicitudEsp"].ToString().Length != 0 && rows["DatosSolicitudEsp"].ToString() != null &&
                                rows["CodigoNombredeGrupoEsp"].ToString().Length != 0 && rows["CodigoNombredeGrupoEsp"].ToString() != null &&
                                rows["TipoNovedadEsp"].ToString().Length != 0 && rows["TipoNovedadEsp"].ToString() != null &&
                                rows["InfoEspecificaActualEsp"].ToString().Length != 0 && rows["InfoEspecificaActualEsp"].ToString() != null &&
                                rows["PuestoTrabajoActEsp"].ToString().Length != 0 && rows["PuestoTrabajoActEsp"].ToString() != null &&
                                rows["GrupoActEsp"].ToString().Length != 0 && rows["GrupoActEsp"].ToString() != null &&
                                rows["CargoActEsp"].ToString().Length != 0 && rows["CargoActEsp"].ToString() != null &&
                                rows["ClasificacionActEsp"].ToString().Length != 0 && rows["ClasificacionActEsp"].ToString() != null &&
                                rows["Nivel1ActEsp"].ToString().Length != 0 && rows["Nivel1ActEsp"].ToString() != null &&
                                rows["Nivel2ActEsp"].ToString().Length != 0 && rows["Nivel2ActEsp"].ToString() != null &&
                                rows["Nivel3ActEsp"].ToString().Length != 0 && rows["Nivel3ActEsp"].ToString() != null &&
                                rows["Nivel4ActEsp"].ToString().Length != 0 && rows["Nivel4ActEsp"].ToString() != null &&
                                rows["Nivel5ActEsp"].ToString().Length != 0 && rows["Nivel5ActEsp"].ToString() != null &&
                                rows["Nivel6ActEsp"].ToString().Length != 0 && rows["Nivel6ActEsp"].ToString() != null &&
                                rows["Nivel7ActEsp"].ToString().Length != 0 && rows["Nivel7ActEsp"].ToString() != null &&
                                rows["InfoEspecificaCambioEsp"].ToString().Length != 0 && rows["InfoEspecificaCambioEsp"].ToString() != null &&
                                rows["PuestoTrabajoCambEsp"].ToString().Length != 0 && rows["PuestoTrabajoCambEsp"].ToString() != null &&
                                rows["GrupoCambEsp"].ToString().Length != 0 && rows["GrupoCambEsp"].ToString() != null &&
                                rows["DiferenciaSalarioEsp"].ToString().Length != 0 && rows["DiferenciaSalarioEsp"].ToString() != null &&
                                rows["ClasificacionCambEsp"].ToString().Length != 0 && rows["ClasificacionCambEsp"].ToString() != null &&
                                rows["Nivel1CambEsp"].ToString().Length != 0 && rows["Nivel1CambEsp"].ToString() != null &&
                                rows["Nivel2CambEsp"].ToString().Length != 0 && rows["Nivel2CambEsp"].ToString() != null &&
                                rows["Nivel3CambEsp"].ToString().Length != 0 && rows["Nivel3CambEsp"].ToString() != null &&
                                rows["Nivel4CambEsp"].ToString().Length != 0 && rows["Nivel4CambEsp"].ToString() != null &&
                                rows["Nivel5CambEsp"].ToString().Length != 0 && rows["Nivel5CambEsp"].ToString() != null &&
                                rows["Nivel6CambEsp"].ToString().Length != 0 && rows["Nivel6CambEsp"].ToString() != null &&
                                rows["Nivel7CambEsp"].ToString().Length != 0 && rows["Nivel7CambEsp"].ToString() != null &&

                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                //string ActualizarEsp = rows["ActualizarEsp"].ToString();
                                string NumSolicitudEsp = rows["NumSolicitudEsp"].ToString();
                                string EstadoEsp = rows["EstadoEsp"].ToString();
                                string EmpresaEsp = rows["EmpresaEsp"].ToString();

                                string DatosColaboradorEsp = rows["DatosColaboradorEsp"].ToString();
                                string IdentificacionColabEsp = rows["IdentificacionColabEsp"].ToString();
                                string NombresColabEsp = rows["NombresColabEsp"].ToString();
                                string ApellidosColabEsp = rows["ApellidosColabEsp"].ToString();
                                string ResponsableSolicitudEsp = rows["ResponsableSolicitudEsp"].ToString();
                                string IdentificacionRespEsp = rows["IdentificacionRespEsp"].ToString();
                                string NombreRespEsp = rows["NombreRespEsp"].ToString();
                                string ApellidosRespEsp = rows["ApellidosRespEsp"].ToString();
                                string DatosSolicitudEsp = rows["DatosSolicitudEsp"].ToString();
                                string CodigoNombredeGrupoEsp = rows["CodigoNombredeGrupoEsp"].ToString();
                                string TipoNovedadEsp = rows["TipoNovedadEsp"].ToString();
                                string InfoEspecificaActualEsp = rows["InfoEspecificaActualEsp"].ToString();
                                string PuestoTrabajoActEsp = rows["PuestoTrabajoActEsp"].ToString();
                                string GrupoActEsp = rows["GrupoActEsp"].ToString();
                                string CargoActEsp = rows["CargoActEsp"].ToString();
                                string ClasificacionActEsp = rows["ClasificacionActEsp"].ToString();
                                string Nivel1ActEsp = rows["Nivel1ActEsp"].ToString();
                                string Nivel2ActEsp = rows["Nivel2ActEsp"].ToString();
                                string Nivel3ActEsp = rows["Nivel3ActEsp"].ToString();
                                string Nivel4ActEsp = rows["Nivel4ActEsp"].ToString();
                                string Nivel5ActEsp = rows["Nivel5ActEsp"].ToString();
                                string Nivel6ActEsp = rows["Nivel6ActEsp"].ToString();
                                string Nivel7ActEsp = rows["Nivel7ActEsp"].ToString();
                                string InfoEspecificaCambioEsp = rows["InfoEspecificaCambioEsp"].ToString();
                                string PuestoTrabajoCambEsp = rows["PuestoTrabajoCambEsp"].ToString();
                                string GrupoCambEsp = rows["GrupoCambEsp"].ToString();
                                string DiferenciaSalarioEsp = rows["DiferenciaSalarioEsp"].ToString();
                                string ClasificacionCambEsp = rows["ClasificacionCambEsp"].ToString();
                                string Nivel1CambEsp = rows["Nivel1CambEsp"].ToString();
                                string Nivel2CambEsp = rows["Nivel2CambEsp"].ToString();
                                string Nivel3CambEsp = rows["Nivel3CambEsp"].ToString();
                                string Nivel4CambEsp = rows["Nivel4CambEsp"].ToString();
                                string Nivel5CambEsp = rows["Nivel5CambEsp"].ToString();
                                string Nivel6CambEsp = rows["Nivel6CambEsp"].ToString();
                                string Nivel7CambEsp = rows["Nivel7CambEsp"].ToString();
                                string url2 = rows["url2"].ToString();


                                List<string> variablesEsperadas = new List<string>() { DatosColaboradorEsp, IdentificacionColabEsp, NombresColabEsp, ApellidosColabEsp, ResponsableSolicitudEsp, IdentificacionRespEsp, NombreRespEsp, ApellidosRespEsp, DatosSolicitudEsp, CodigoNombredeGrupoEsp, TipoNovedadEsp, InfoEspecificaActualEsp, PuestoTrabajoActEsp, GrupoActEsp, CargoActEsp, ClasificacionActEsp, Nivel1ActEsp, Nivel2ActEsp, Nivel3ActEsp, Nivel4ActEsp, Nivel5ActEsp, Nivel6ActEsp, Nivel7ActEsp, InfoEspecificaCambioEsp, PuestoTrabajoCambEsp, GrupoCambEsp, DiferenciaSalarioEsp, ClasificacionCambEsp, Nivel1CambEsp, Nivel2CambEsp, Nivel3CambEsp, Nivel4CambEsp, Nivel5CambEsp, Nivel6CambEsp, Nivel7CambEsp };


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
                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://ophtsph:8085/".ToLower())
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

                                    //Clic en RolLider
                                    //Rol Lider
                                    if (database == "ORA")
                                    {
                                        selenium.Click("//*[@id='pLider']");
                                    }
                                    else
                                    {
                                        selenium.Click("//span[contains(.,'Rol Lider')]");
                                    }
                                    Thread.Sleep(6000);

                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);

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

                                    selenium.Screenshot("Titulos Descriptivos", true, file);

                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);


                                    //Validación Emergentes campos

                                    string NumSol = selenium.CamposEmergentes("//span[@id='ctl00_ContenidoPagina_lblsoli']", "Número de Solicitud", file);
                                    Thread.Sleep(100);
                                    if (NumSol != NumSolicitudEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + NumSolicitudEsp + " y el encontrado es: " + NumSol);
                                    }
                                    Thread.Sleep(100);

                                    string Estado = selenium.CamposEmergentes("//span[@id='ctl00_ContenidoPagina_lblEsta']", "Estado", file);
                                    Thread.Sleep(100);
                                    if (Estado != EstadoEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + EstadoEsp + " y el encontrado es: " + Estado);
                                    }
                                    Thread.Sleep(100);

                                    string Empresa = selenium.CamposEmergentes("//span[@id='ctl00_ContenidoPagina_lblEmpresa']", "Empresa", file);
                                    Thread.Sleep(100);
                                    if (Empresa != EmpresaEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + EmpresaEsp + " y el encontrado es: " + Empresa);
                                    }
                                    Thread.Sleep(100);

                                    //Clic en Gestionar Puesto
                                    selenium.Click("//a[contains(text(),'Gestionar puesto')]");
                                    Thread.Sleep(100);

                                    //LISTA XPATH
                                    List<string> xpath = new List<string>() {
                                        "//span[@id='ctl00_ContenidoPagina_Label7']",
                                        "//span[@id='ctl00_ContenidoPagina_lblTelResi']",
                                        "//span[@id='ctl00_ContenidoPagina_Label1']",
                                        "//span[@id='ctl00_ContenidoPagina_txtApe_empl']",
                                        "//span[@id='ctl00_ContenidoPagina_Label8']",
                                        "//span[@id='ctl00_ContenidoPagina_Label2']",
                                        "//span[@id='ctl00_ContenidoPagina_Label3']",
                                        "//span[@id='ctl00_ContenidoPagina_Label4']",
                                        "//span[@id='ctl00_ContenidoPagina_Label9']",
                                        "//span[contains(@id,'ctl00_ContenidoPagina_Label5')]",
                                        "//span[@id='ctl00_ContenidoPagina_lblTipNove']",
                                        "//span[@id='ctl00_ContenidoPagina_lblInfActu']",
                                        "//span[@id='ctl00_ContenidoPagina_lblPueTraA']",
                                        "//span[@id='ctl00_ContenidoPagina_lblGrupo']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCargo']",
                                        "//span[@id='ctl00_ContenidoPagina_Label12']",
                                        "//span[@id='ctl00_ContenidoPagina_Label19']",
                                        "//span[@id='ctl00_ContenidoPagina_Label20']",
                                        "//span[@id='ctl00_ContenidoPagina_Label21']",
                                        "//span[@id='ctl00_ContenidoPagina_Label22']",
                                        "//span[@id='ctl00_ContenidoPagina_Label23']",
                                        "//span[@id='ctl00_ContenidoPagina_Label24']",
                                        "//span[@id='ctl00_ContenidoPagina_Label25']",
                                        "//span[@id='ctl00_ContenidoPagina_Label10']",
                                        "//span[@id='ctl00_ContenidoPagina_lblPueTraC']",
                                        "//span[@id='ctl00_ContenidoPagina_lblGrupoC']",
                                        "//span[@id='ctl00_ContenidoPagina_lblDifSalalbl']",
                                        "//span[@id='ctl00_ContenidoPagina_Label11']",
                                        "//span[@id='ctl00_ContenidoPagina_lblNiv1']",
                                        "//span[@id='ctl00_ContenidoPagina_Label13']",
                                        "//span[@id='ctl00_ContenidoPagina_Label14']",
                                        "//span[@id='ctl00_ContenidoPagina_Label15']",
                                        "//span[@id='ctl00_ContenidoPagina_Label16']",
                                        "//span[@id='ctl00_ContenidoPagina_Label17']",
                                        "//span[@id='ctl00_ContenidoPagina_Label18']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSC = new List<string>() {
                                        "Datos Colaborador", "Identificación Colaborador", "Nombres Colaborador", "Apellidos Colaborador",
                                        "Responsable Solicitud","Identificación Responsable","Nombres Responsable",
                                        "Apellidos Responsable","Datos de la Solicitud","Código / Nombre del grupo","Tipo de Novedad",
                                        "Información Específica Actual","Puesto de Trabajo Actual","Grupo Actual","Cargo Actual", "Clasificación Actual",
                                        "Nivel 1 Actual", "Nivel 2 Actual", "Nivel 3 Actual", "Nivel 4 Actual", "Nivel 5 Actual", "Nivel 6 Actual",
                                        "Nivel 7 Actual", "Información Específica Cambio","Puesto de Trabajo Cambio","Grupo Cambio",
                                        "Diferencia Salario Cambio","Clasificación Cambio","Nivel 1 Cambio", "Nivel 2 Cambio", "Nivel 3 Cambio", "Nivel 4 Cambio",
                                        "Nivel 5 Cambio", "Nivel 6 Cambio", "Nivel 7 Cambio"
                                    };

                                    //SCROLL
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_navgestionar']/div/div");
                                    for (int j = 0; j < 9; j++)
                                    {
                                        Keyboard.SendKeys("{DOWN}");
                                    }

                                    //Validación Emergentes campos
                                    for (int i = 0; i < 35; i++)
                                    {
                                        if (i != 9)
                                        {
                                            string campoName = selenium.CamposEmergentes(xpath[i], descripcionSC[i], file);
                                            Thread.Sleep(100);
                                            if (campoName.Trim() != variablesEsperadas[i].Trim())
                                            {
                                                errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas[i] + " y el encontrado es: " + campoName);
                                            }
                                            Thread.Sleep(100);
                                            if (i == 10)
                                            {
                                                for (int j = 0; j < 8; j++)
                                                {
                                                    Keyboard.SendKeys("{DOWN}");
                                                }
                                            }
                                        }
                                    }

                                    //SCROLL
                                    for (int k = 0; k < 17; k++)
                                    {
                                        Keyboard.SendKeys("{UP}");
                                    }


                                    ////////////// VALIDACION DE TABS/////////////////

                                    selenium.ValTabs(file);

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
        public void SmartPeople_frmBiDatadNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_6.SmartPeople_frmBiDatadNTC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                rows["ActualizarEsp"].ToString().Length != 0 && rows["ActualizarEsp"].ToString() != null &&
                                //Otros datos
                                rows["MisHobiesEsp"].ToString().Length != 0 && rows["MisHobiesEsp"].ToString() != null &&
                                rows["MisPropositosEsp"].ToString().Length != 0 && rows["MisPropositosEsp"].ToString() != null &&
                                rows["ObservacionesEsp"].ToString().Length != 0 && rows["ObservacionesEsp"].ToString() != null &&
                                rows["ProfesionuOficioEsp"].ToString().Length != 0 && rows["ProfesionuOficioEsp"].ToString() != null &&
                                rows["AñosEsp"].ToString().Length != 0 && rows["AñosEsp"].ToString() != null &&
                                rows["IngresosAdicionalesEsp"].ToString().Length != 0 && rows["IngresosAdicionalesEsp"].ToString() != null &&
                                rows["RegionesPaisesConocidosEsp"].ToString().Length != 0 && rows["RegionesPaisesConocidosEsp"].ToString() != null &&
                                rows["NumeroCajaEsp"].ToString().Length != 0 && rows["NumeroCajaEsp"].ToString() != null &&
                                rows["PosicionCajaEsp"].ToString().Length != 0 && rows["PosicionCajaEsp"].ToString() != null &&
                                rows["JornadaLaboralEsp"].ToString().Length != 0 && rows["JornadaLaboralEsp"].ToString() != null &&
                                rows["MascotasEsp"].ToString().Length != 0 && rows["MascotasEsp"].ToString() != null &&
                                rows["LateralidadEsp"].ToString().Length != 0 && rows["LateralidadEsp"].ToString() != null &&
                                //Salud
                                rows["GrupoSangreEsp"].ToString().Length != 0 && rows["GrupoSangreEsp"].ToString() != null &&
                                rows["FactorEsp"].ToString().Length != 0 && rows["FactorEsp"].ToString() != null &&
                                rows["EstaturaEsp"].ToString().Length != 0 && rows["EstaturaEsp"].ToString() != null &&
                                rows["PesoEsp"].ToString().Length != 0 && rows["PesoEsp"].ToString() != null &&
                                rows["RazaEsp"].ToString().Length != 0 && rows["RazaEsp"].ToString() != null &&
                                rows["MedicoPersonalEsp"].ToString().Length != 0 && rows["MedicoPersonalEsp"].ToString() != null &&
                                rows["TelefonoEsp"].ToString().Length != 0 && rows["TelefonoEsp"].ToString() != null &&
                                rows["EmpresaSegEsp"].ToString().Length != 0 && rows["EmpresaSegEsp"].ToString() != null &&
                                rows["NumeroSegEsp"].ToString().Length != 0 && rows["NumeroSegEsp"].ToString() != null &&
                                rows["EnfermedadesPadeceEsp"].ToString().Length != 0 && rows["EnfermedadesPadeceEsp"].ToString() != null &&
                                rows["RestriccionMedicaEsp"].ToString().Length != 0 && rows["RestriccionMedicaEsp"].ToString() != null &&
                                rows["AlergicoEsp"].ToString().Length != 0 && rows["AlergicoEsp"].ToString() != null &&
                                rows["LlamarEmergenciaEsp"].ToString().Length != 0 && rows["LlamarEmergenciaEsp"].ToString() != null &&
                                rows["TelefonoContactoEsp"].ToString().Length != 0 && rows["TelefonoContactoEsp"].ToString() != null &&
                                //Servicio salud
                                rows["PlanServSaludEsp"].ToString().Length != 0 && rows["PlanServSaludEsp"].ToString() != null &&
                                rows["OtroPlanServSaludEsp"].ToString().Length != 0 && rows["OtroPlanServSaludEsp"].ToString() != null &&
                                rows["EntidadEsp"].ToString().Length != 0 && rows["EntidadEsp"].ToString() != null &&
                                rows["GrupoApoyoEsp"].ToString().Length != 0 && rows["GrupoApoyoEsp"].ToString() != null &&
                                rows["CualServSaludEsp"].ToString().Length != 0 && rows["CualServSaludEsp"].ToString() != null &&
                                rows["EnfermedadLaboralEsp"].ToString().Length != 0 && rows["EnfermedadLaboralEsp"].ToString() != null &&
                                rows["GradoPerdidaSaludEsp"].ToString().Length != 0 && rows["GradoPerdidaSaludEsp"].ToString() != null &&
                                rows["RestriccionMedicaServSaludEsp"].ToString().Length != 0 && rows["RestriccionMedicaServSaludEsp"].ToString() != null &&
                                rows["EntidadEmitioEsp"].ToString().Length != 0 && rows["EntidadEmitioEsp"].ToString() != null &&
                                //Vivienda                                
                                rows["TipoViviendaEsp"].ToString().Length != 0 && rows["TipoViviendaEsp"].ToString() != null &&
                                rows["ClaseEsp"].ToString().Length != 0 && rows["ClaseEsp"].ToString() != null &&
                                rows["EspecifiqueEsp"].ToString().Length != 0 && rows["EspecifiqueEsp"].ToString() != null &&
                                rows["NombreArrendadorEsp"].ToString().Length != 0 && rows["NombreArrendadorEsp"].ToString() != null &&
                                rows["TelefonoViviendaEsp"].ToString().Length != 0 && rows["TelefonoViviendaEsp"].ToString() != null &&
                                rows["PerimetroViviendaEsp"].ToString().Length != 0 && rows["PerimetroViviendaEsp"].ToString() != null &&
                                rows["EstratoViviendaEsp"].ToString().Length != 0 && rows["EstratoViviendaEsp"].ToString() != null &&
                                rows["NumHabitacionesEsp"].ToString().Length != 0 && rows["NumHabitacionesEsp"].ToString() != null &&
                                rows["NumPersonasHabitanEsp"].ToString().Length != 0 && rows["NumPersonasHabitanEsp"].ToString() != null &&
                                rows["ServiciosViviendaEsp"].ToString().Length != 0 && rows["ServiciosViviendaEsp"].ToString() != null &&
                                rows["BeneficiarioCreditoEsp"].ToString().Length != 0 && rows["BeneficiarioCreditoEsp"].ToString() != null &&
                                rows["CreditoViviendaEsp"].ToString().Length != 0 && rows["CreditoViviendaEsp"].ToString() != null &&
                                rows["HabitantesViviendaEsp"].ToString().Length != 0 && rows["HabitantesViviendaEsp"].ToString() != null &&
                                rows["IdentificacionViviendaEsp"].ToString().Length != 0 && rows["IdentificacionViviendaEsp"].ToString() != null &&
                                rows["PlanComplementarioEsp"].ToString().Length != 0 && rows["PlanComplementarioEsp"].ToString() != null &&
                                rows["PlanMedicinaEsp"].ToString().Length != 0 && rows["PlanMedicinaEsp"].ToString() != null &&
                                //TIempo libre
                                rows["OficiosDomesticosEsp"].ToString().Length != 0 && rows["OficiosDomesticosEsp"].ToString() != null &&
                                rows["TiempoTareaEsp"].ToString().Length != 0 && rows["TiempoTareaEsp"].ToString() != null &&
                                rows["EspecifiqueOfiDomesEsp"].ToString().Length != 0 && rows["EspecifiqueOfiDomesEsp"].ToString() != null &&
                                rows["RecreacionEsp"].ToString().Length != 0 && rows["RecreacionEsp"].ToString() != null &&
                                rows["PeriocidadRecreacionEsp"].ToString().Length != 0 && rows["PeriocidadRecreacionEsp"].ToString() != null &&
                                rows["EspecifiqueRecreacionEsp"].ToString().Length != 0 && rows["EspecifiqueRecreacionEsp"].ToString() != null &&
                                rows["DeporteEsp"].ToString().Length != 0 && rows["DeporteEsp"].ToString() != null &&
                                rows["PeriocidadDeporteEsp"].ToString().Length != 0 && rows["PeriocidadDeporteEsp"].ToString() != null &&
                                rows["EspecifiqueDeporteEsp"].ToString().Length != 0 && rows["EspecifiqueDeporteEsp"].ToString() != null &&
                                rows["OtroTrabajoEsp"].ToString().Length != 0 && rows["OtroTrabajoEsp"].ToString() != null &&
                                rows["PeriocidadOtroTrabajoEsp"].ToString().Length != 0 && rows["PeriocidadOtroTrabajoEsp"].ToString() != null &&
                                rows["EspecifiqueOtroTrabajoEsp"].ToString().Length != 0 && rows["EspecifiqueOtroTrabajoEsp"].ToString() != null &&
                                //Medio desplazamiento
                                rows["VehiculoPropioEsp"].ToString().Length != 0 && rows["VehiculoPropioEsp"].ToString() != null &&
                                rows["PlacaVehiculoEsp"].ToString().Length != 0 && rows["PlacaVehiculoEsp"].ToString() != null &&
                                rows["LicenciaConduccionEsp"].ToString().Length != 0 && rows["LicenciaConduccionEsp"].ToString() != null &&
                                rows["TransportePublicoEsp"].ToString().Length != 0 && rows["TransportePublicoEsp"].ToString() != null &&
                                rows["BusesporRecorridoEsp"].ToString().Length != 0 && rows["BusesporRecorridoEsp"].ToString() != null &&
                                rows["EspecifiqueDesplaEsp"].ToString().Length != 0 && rows["EspecifiqueDesplaEsp"].ToString() != null &&
                                rows["TransporteDesplaEsp"].ToString().Length != 0 && rows["TransporteDesplaEsp"].ToString() != null &&
                                rows["FrecuenciaTrasladosEsp"].ToString().Length != 0 && rows["FrecuenciaTrasladosEsp"].ToString() != null &&
                                rows["MedioTransporteEsp"].ToString().Length != 0 && rows["MedioTransporteEsp"].ToString() != null &&
                                rows["TiempoDesplazamientoEsp"].ToString().Length != 0 && rows["TiempoDesplazamientoEsp"].ToString() != null &&
                                //Oficina
                                rows["DireccionOficinaEsp"].ToString().Length != 0 && rows["DireccionOficinaEsp"].ToString() != null &&
                                rows["PaisOficinaEsp"].ToString().Length != 0 && rows["PaisOficinaEsp"].ToString() != null &&
                                rows["DepartamentoOficinaEsp"].ToString().Length != 0 && rows["DepartamentoOficinaEsp"].ToString() != null &&
                                rows["MunicipioOficinaEsp"].ToString().Length != 0 && rows["MunicipioOficinaEsp"].ToString() != null &&
                                rows["TelefonoOficinaEsp"].ToString().Length != 0 && rows["TelefonoOficinaEsp"].ToString() != null &&
                                rows["ExtensionOficinaEsp"].ToString().Length != 0 && rows["ExtensionOficinaEsp"].ToString() != null &&
                                rows["FaxOficinaEsp"].ToString().Length != 0 && rows["FaxOficinaEsp"].ToString() != null &&
                                //Datos complementarios
                                rows["DependePerDiscapacitadasEsp"].ToString().Length != 0 && rows["DependePerDiscapacitadasEsp"].ToString() != null &&
                                rows["ParentescoDiscapaEsp"].ToString().Length != 0 && rows["ParentescoDiscapaEsp"].ToString() != null &&
                                rows["PerteneceGrupoSocialEsp"].ToString().Length != 0 && rows["PerteneceGrupoSocialEsp"].ToString() != null &&
                                rows["CualGrupoSocialEsp"].ToString().Length != 0 && rows["CualGrupoSocialEsp"].ToString() != null &&
                                rows["PeriocidadReuneEsp"].ToString().Length != 0 && rows["PeriocidadReuneEsp"].ToString() != null &&
                                rows["AporteGrupoEsp"].ToString().Length != 0 && rows["AporteGrupoEsp"].ToString() != null &&
                                rows["CondicionEspecialEsp"].ToString().Length != 0 && rows["CondicionEspecialEsp"].ToString() != null &&
                                rows["CualCondicionEspecialEsp"].ToString().Length != 0 && rows["CualCondicionEspecialEsp"].ToString() != null &&
                                rows["ComunidadLGBTEsp"].ToString().Length != 0 && rows["ComunidadLGBTEsp"].ToString() != null &&
                                //Url2
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                string ActualizarEsp = rows["ActualizarEsp"].ToString();
                                //Otros datos
                                string MisHobiesEsp = rows["MisHobiesEsp"].ToString();
                                string MisPropositosEsp = rows["MisPropositosEsp"].ToString();
                                string ObservacionesEsp = rows["ObservacionesEsp"].ToString();
                                string ProfesionuOficioEsp = rows["ProfesionuOficioEsp"].ToString();
                                string AñosEsp = rows["AñosEsp"].ToString();
                                string IngresosAdicionalesEsp = rows["IngresosAdicionalesEsp"].ToString();
                                string RegionesPaisesConocidosEsp = rows["RegionesPaisesConocidosEsp"].ToString();
                                string NumeroCajaEsp = rows["NumeroCajaEsp"].ToString();
                                string PosicionCajaEsp = rows["PosicionCajaEsp"].ToString();
                                string JornadaLaboralEsp = rows["JornadaLaboralEsp"].ToString();
                                string MascotasEsp = rows["MascotasEsp"].ToString();
                                string LateralidadEsp = rows["LateralidadEsp"].ToString();
                                //Salud
                                string GrupoSangreEsp = rows["GrupoSangreEsp"].ToString();
                                string FactorEsp = rows["FactorEsp"].ToString();
                                string EstaturaEsp = rows["EstaturaEsp"].ToString();
                                string PesoEsp = rows["PesoEsp"].ToString();
                                string RazaEsp = rows["RazaEsp"].ToString();
                                string MedicoPersonalEsp = rows["MedicoPersonalEsp"].ToString();
                                string TelefonoEsp = rows["TelefonoEsp"].ToString();
                                string EmpresaSegEsp = rows["EmpresaSegEsp"].ToString();
                                string NumeroSegEsp = rows["NumeroSegEsp"].ToString();
                                string EnfermedadesPadeceEsp = rows["EnfermedadesPadeceEsp"].ToString();
                                string RestriccionMedicaEsp = rows["RestriccionMedicaEsp"].ToString();
                                string AlergicoEsp = rows["AlergicoEsp"].ToString();
                                string LlamarEmergenciaEsp = rows["LlamarEmergenciaEsp"].ToString();
                                string TelefonoContactoEsp = rows["TelefonoContactoEsp"].ToString();
                                //Servicio salud
                                string PlanServSaludEsp = rows["PlanServSaludEsp"].ToString();
                                string OtroPlanServSaludEsp = rows["OtroPlanServSaludEsp"].ToString();
                                string EntidadEsp = rows["EntidadEsp"].ToString();
                                string GrupoApoyoEsp = rows["GrupoApoyoEsp"].ToString();
                                string CualServSaludEsp = rows["CualServSaludEsp"].ToString();
                                string EnfermedadLaboralEsp = rows["EnfermedadLaboralEsp"].ToString();
                                string GradoPerdidaSaludEsp = rows["GradoPerdidaSaludEsp"].ToString();
                                string RestriccionMedicaServSaludEsp = rows["RestriccionMedicaServSaludEsp"].ToString();
                                string EntidadEmitioEsp = rows["EntidadEmitioEsp"].ToString();
                                //Vivienda
                                string TipoViviendaEsp = rows["TipoViviendaEsp"].ToString();
                                string ClaseEsp = rows["ClaseEsp"].ToString();
                                string EspecifiqueEsp = rows["EspecifiqueEsp"].ToString();
                                string NombreArrendadorEsp = rows["NombreArrendadorEsp"].ToString();
                                string TelefonoViviendaEsp = rows["TelefonoViviendaEsp"].ToString();
                                string PerimetroViviendaEsp = rows["PerimetroViviendaEsp"].ToString();
                                string EstratoViviendaEsp = rows["EstratoViviendaEsp"].ToString();
                                string NumHabitacionesEsp = rows["NumHabitacionesEsp"].ToString();
                                string NumPersonasHabitanEsp = rows["NumPersonasHabitanEsp"].ToString();
                                string ServiciosViviendaEsp = rows["ServiciosViviendaEsp"].ToString();
                                string BeneficiarioCreditoEsp = rows["BeneficiarioCreditoEsp"].ToString();
                                string CreditoViviendaEsp = rows["CreditoViviendaEsp"].ToString();
                                string HabitantesViviendaEsp = rows["HabitantesViviendaEsp"].ToString();
                                string IdentificacionViviendaEsp = rows["IdentificacionViviendaEsp"].ToString();
                                string PlanComplementarioEsp = rows["PlanComplementarioEsp"].ToString();
                                string PlanMedicinaEsp = rows["PlanMedicinaEsp"].ToString();
                                //TIempo libre
                                string OficiosDomesticosEsp = rows["OficiosDomesticosEsp"].ToString();
                                string TiempoTareaEsp = rows["TiempoTareaEsp"].ToString();
                                string EspecifiqueOfiDomesEsp = rows["EspecifiqueOfiDomesEsp"].ToString();
                                string RecreacionEsp = rows["RecreacionEsp"].ToString();
                                string PeriocidadRecreacionEsp = rows["PeriocidadRecreacionEsp"].ToString();
                                string EspecifiqueRecreacionEsp = rows["EspecifiqueRecreacionEsp"].ToString();
                                string DeporteEsp = rows["DeporteEsp"].ToString();
                                string PeriocidadDeporteEsp = rows["PeriocidadDeporteEsp"].ToString();
                                string EspecifiqueDeporteEsp = rows["EspecifiqueDeporteEsp"].ToString();
                                string OtroTrabajoEsp = rows["OtroTrabajoEsp"].ToString();
                                string PeriocidadOtroTrabajoEsp = rows["PeriocidadOtroTrabajoEsp"].ToString();
                                string EspecifiqueOtroTrabajoEsp = rows["EspecifiqueOtroTrabajoEsp"].ToString();
                                //Medio desplazamiento
                                string VehiculoPropioEsp = rows["VehiculoPropioEsp"].ToString();
                                string PlacaVehiculoEsp = rows["PlacaVehiculoEsp"].ToString();
                                string LicenciaConduccionEsp = rows["LicenciaConduccionEsp"].ToString();
                                string TransportePublicoEsp = rows["TransportePublicoEsp"].ToString();
                                string BusesporRecorridoEsp = rows["BusesporRecorridoEsp"].ToString();
                                string EspecifiqueDesplaEsp = rows["EspecifiqueDesplaEsp"].ToString();
                                string TransporteDesplaEsp = rows["TransporteDesplaEsp"].ToString();
                                string FrecuenciaTrasladosEsp = rows["FrecuenciaTrasladosEsp"].ToString();
                                string MedioTransporteEsp = rows["MedioTransporteEsp"].ToString();
                                string TiempoDesplazamientoEsp = rows["TiempoDesplazamientoEsp"].ToString();
                                //Oficina
                                string DireccionOficinaEsp = rows["DireccionOficinaEsp"].ToString();
                                string PaisOficinaEsp = rows["PaisOficinaEsp"].ToString();
                                string DepartamentoOficinaEsp = rows["DepartamentoOficinaEsp"].ToString();
                                string MunicipioOficinaEsp = rows["MunicipioOficinaEsp"].ToString();
                                string TelefonoOficinaEsp = rows["TelefonoOficinaEsp"].ToString();
                                string ExtensionOficinaEsp = rows["ExtensionOficinaEsp"].ToString();
                                string FaxOficinaEsp = rows["FaxOficinaEsp"].ToString();
                                //Datos complementarios
                                string DependePerDiscapacitadasEsp = rows["DependePerDiscapacitadasEsp"].ToString();
                                string ParentescoDiscapaEsp = rows["ParentescoDiscapaEsp"].ToString();
                                string PerteneceGrupoSocialEsp = rows["PerteneceGrupoSocialEsp"].ToString();
                                string CualGrupoSocialEsp = rows["CualGrupoSocialEsp"].ToString();
                                string PeriocidadReuneEsp = rows["PeriocidadReuneEsp"].ToString();
                                string AporteGrupoEsp = rows["AporteGrupoEsp"].ToString();
                                string CondicionEspecialEsp = rows["CondicionEspecialEsp"].ToString();
                                string CualCondicionEspecialEsp = rows["CualCondicionEspecialEsp"].ToString();
                                string ComunidadLGBTEsp = rows["ComunidadLGBTEsp"].ToString();
                                //Url2
                                string url2 = rows["url2"].ToString();

                                List<string> variablesEsperadas = new List<string>() { MisHobiesEsp, MisPropositosEsp, ObservacionesEsp, ProfesionuOficioEsp,
                                    AñosEsp, IngresosAdicionalesEsp, RegionesPaisesConocidosEsp, NumeroCajaEsp, PosicionCajaEsp, JornadaLaboralEsp,
                                    MascotasEsp, LateralidadEsp };

                                List<string> variablesEsperadas1 = new List<string>() { GrupoSangreEsp, FactorEsp, EstaturaEsp, PesoEsp, RazaEsp,
                                    MedicoPersonalEsp, TelefonoEsp, EmpresaSegEsp, NumeroSegEsp, EnfermedadesPadeceEsp, RestriccionMedicaEsp,
                                    AlergicoEsp, LlamarEmergenciaEsp, TelefonoContactoEsp };

                                List<string> variablesEsperadas2 = new List<string>() { PlanServSaludEsp, OtroPlanServSaludEsp, EntidadEsp, GrupoApoyoEsp,
                                    CualServSaludEsp, EnfermedadLaboralEsp, GradoPerdidaSaludEsp, RestriccionMedicaServSaludEsp, EntidadEmitioEsp };

                                List<string> variablesEsperadas3 = new List<string>() { TipoViviendaEsp, ClaseEsp, EspecifiqueEsp, NombreArrendadorEsp, TelefonoViviendaEsp,
                                    PerimetroViviendaEsp, EstratoViviendaEsp, NumHabitacionesEsp, NumPersonasHabitanEsp, ServiciosViviendaEsp,
                                    BeneficiarioCreditoEsp, CreditoViviendaEsp, HabitantesViviendaEsp, IdentificacionViviendaEsp, PlanComplementarioEsp, PlanMedicinaEsp };

                                List<string> variablesEsperadas4 = new List<string>() { OficiosDomesticosEsp, TiempoTareaEsp, EspecifiqueOfiDomesEsp,
                                    RecreacionEsp, PeriocidadRecreacionEsp, EspecifiqueRecreacionEsp, DeporteEsp, PeriocidadDeporteEsp,
                                    EspecifiqueDeporteEsp, OtroTrabajoEsp, PeriocidadOtroTrabajoEsp, EspecifiqueOtroTrabajoEsp  };

                                List<string> variablesEsperadas5 = new List<string>() { VehiculoPropioEsp, PlacaVehiculoEsp, LicenciaConduccionEsp, TransportePublicoEsp,
                                    BusesporRecorridoEsp, EspecifiqueDesplaEsp, TransporteDesplaEsp, FrecuenciaTrasladosEsp, MedioTransporteEsp, TiempoDesplazamientoEsp };

                                List<string> variablesEsperadas6 = new List<string>() {DireccionOficinaEsp, PaisOficinaEsp, DepartamentoOficinaEsp, MunicipioOficinaEsp,
                                    TelefonoOficinaEsp, ExtensionOficinaEsp, FaxOficinaEsp };

                                List<string> variablesEsperadas7 = new List<string>() {DependePerDiscapacitadasEsp, ParentescoDiscapaEsp, PerteneceGrupoSocialEsp, CualGrupoSocialEsp,
                                    PeriocidadReuneEsp, AporteGrupoEsp, CondicionEspecialEsp, CualCondicionEspecialEsp, ComunidadLGBTEsp };


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
                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://ophtsph:8085/".ToLower())
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

                                    ChromeDriver driver2 = selenium.returnDriver();

                                    driver2.Navigate().GoToUrl(url2);

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

                                    selenium.Screenshot("Titulos Descriptivos", true, file);

                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Actualizar
                                    string Actualizar = selenium.EmergenteBotones("btnGuardar");
                                    Thread.Sleep(100);
                                    if (Actualizar != ActualizarEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Actualizar es incorrecto, el esperado es: " + ActualizarEsp + " y el encontrado es: " + Actualizar);
                                    }
                                    Thread.Sleep(100);

                                    //////////  VALIDACION EMERGENTE DE CAMPOS  //////////

                                    /// -- OTROS DATOS -- ///

                                    //Clic en Otros Datos
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnOtrDa']/span");
                                    Thread.Sleep(100);

                                    //LISTA XPATH OTROS DATOS
                                    List<string> xpathOtrosDatos = new List<string>() {
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_KCtrlTxtHOB_EMPL_lblTexto']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_KCtrlTxtMIS_PROP_lblTexto']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_txtObsErva_lblTexto']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_lblProFesi']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_lblAnoExpo']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_lblIngrAd']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_lblPaiCono']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_lblNumCaja']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_lblPosCaja']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_lblJorLabo']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_txtMascota_lblTexto']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_lbllateralidad']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionOtrosDatos = new List<string>()
                                    {
                                        "Mis Hobbies", "Mis Propósitos", "Observaciones", "Profesión u Oficio", "Años", "Ingresos adicionales",
                                        "Regiones / Países Conocidos", "Número de Caja", "Posición en la Caja", "Jornada Laboral", "Mascotas", "Lateralidad"
                                    };

                                    //SCROLL
                                    //selenium.Click("//div[@id='ctl00_ContenidoPagina_navgestionar']/div/div");
                                    for (int j = 0; j < 2; j++)
                                    {
                                        Keyboard.SendKeys("{DOWN}");
                                    }

                                    //Validación Emergentes campos
                                    for (int i = 0; i < variablesEsperadas.Count; i++)
                                    {
                                        if (i != 6)
                                        {
                                            string campoName = selenium.CamposEmergentes(xpathOtrosDatos[i], descripcionOtrosDatos[i], file);
                                            Thread.Sleep(100);
                                            if (campoName.Trim() != variablesEsperadas[i].Trim())
                                            {
                                                errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas[i] + " y el encontrado es: " + campoName);
                                            }
                                            Thread.Sleep(100);
                                        }
                                        if (i == 8)
                                        {
                                            for (int j = 0; j < 2; j++)
                                            {
                                                Keyboard.SendKeys("{DOWN}");
                                            }
                                        }

                                    }

                                    //SCROLL
                                    //selenium.Click("//div[@id='ctl00_ContenidoPagina_navgestionar']/div/div");
                                    for (int j = 0; j < 4; j++)
                                    {
                                        Keyboard.SendKeys("{UP}");
                                    }

                                    /// -- SALUD -- ///

                                    //Clic en Salud
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnSalud']/span");
                                    Thread.Sleep(100);

                                    //LISTA XPATH SALUD
                                    List<string> xpathSalud = new List<string>() {
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_lblGruSang']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_lblFacSang']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_lblEstEmpl']']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_lblEstPeso']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_lblCodRaza']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_lblMedPers']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_lblTelMedi']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_lblEmpSegu']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_lblNumSegu']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtEnfAler_lblTexto']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtResMedi_lblTexto']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_lblAleRgic']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_lblCasEmer']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_lblTelEmer']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionSalud = new List<string>()
                                    {
                                        "Grupo Sangre", "Factor", "Estatura", "Peso", "Raza", "Médico Personal", "Teléfono", "Empresa Seguro", "Número Seguro",
                                        "Enfermedades que Padece", "Restricción Médica", "Alérgico", "En Caso de Emergencia Llamar", "Teléfono Contacto"
                                    };

                                    //SCROLL
                                    //selenium.Click("//div[@id='ctl00_ContenidoPagina_navgestionar']/div/div");
                                    for (int k = 0; k < 3; k++)
                                    {
                                        Keyboard.SendKeys("{DOWN}");
                                    }

                                    //Validación Emergentes campos
                                    for (int q = 0; q < 11; q++)
                                    {
                                        if (q != 2)
                                        {
                                            string campoName1 = selenium.CamposEmergentes(xpathSalud[q], descripcionSalud[q], file);
                                            Console.WriteLine(campoName1);
                                            Thread.Sleep(100);
                                            if (campoName1.Trim() != variablesEsperadas1[q].Trim())
                                            {
                                                errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas1[q] + " y el encontrado es: " + campoName1);
                                            }
                                            Thread.Sleep(100);

                                        }
                                    }

                                    Thread.Sleep(500);
                                    //SCROLL
                                    for (int k = 0; k < 2; k++)
                                    {
                                        Keyboard.SendKeys("{DOWN}");
                                    }

                                    for (int q = 11; q < variablesEsperadas1.Count; q++)
                                    {
                                        if (q != 2)
                                        {
                                            string campoName1 = selenium.CamposEmergentes(xpathSalud[q], descripcionSalud[q], file);
                                            Console.WriteLine(campoName1);
                                            Thread.Sleep(100);
                                            if (campoName1.Trim() != variablesEsperadas1[q].Trim())
                                            {
                                                errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas1[q] + " y el encontrado es: " + campoName1);
                                            }
                                            Thread.Sleep(100);

                                        }
                                    }


                                    Thread.Sleep(500);
                                    //SCROLL
                                    for (int k = 0; k < 5; k++)
                                    {
                                        Keyboard.SendKeys("{UP}");
                                    }


;                                   /// -- SERVICIO SALUD -- ///

                                    //Clic en Servicio salud
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnSerSalu']/span");
                                    Thread.Sleep(100);

                                    //LISTA XPATH SERVICIO SALUD
                                    List<string> xpathServicioSalud = new List<string>() {
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_lblPlaSalu']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_lblOtrPlan']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_lblNomEnti']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_lblGruApoy']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_lblOtrPoli']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_lblENFLACA']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_lblPorInca']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_KCtrlTxtRES_MEDS_lblTexto']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_lblENT_MEDS']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionServicioSalud = new List<string>()
                                    {
                                        "Plan Servicio", "Otro Plan", "Entidad Servicio", "Grupo Apoyo", "Cual Grupo de Apoyo", "Enfermedad Laboral",
                                        "Grado de Perdida Salud", "Restricción Médica", "Entidad que la emitió"
                                    };


                                    //Validación Emergentes campos
                                    for (int w = 0; w < variablesEsperadas2.Count; w++)
                                    {
                                        string campoName2 = selenium.CamposEmergentes(xpathServicioSalud[w], descripcionServicioSalud[w], file);
                                        Thread.Sleep(100);
                                        if (campoName2.Trim() != variablesEsperadas2[w].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas2[w] + " y el encontrado es: " + campoName2);
                                        }
                                        Thread.Sleep(100);
                                        if (w == 6)
                                        {
                                            for (int j = 0; j < 3; j++)
                                            {
                                                Keyboard.SendKeys("{DOWN}");
                                            }
                                        }
                                    }

                                    Thread.Sleep(500);
                                    //SCROLL
                                    for (int k = 0; k < 5; k++)
                                    {
                                        Keyboard.SendKeys("{UP}");
                                    }

                                    /// -- VIVIENDA -- ///

                                    //Clic en Vivienda
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnVivienda']/span");
                                    Thread.Sleep(100);

                                    //LISTA XPATH SERVICIO VIVIENDA
                                    List<string> xpathVivienda = new List<string>() {
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblVivProp']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblTipVivi']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblTivOtro']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblNomArre']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblTelArre']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblPerVivi']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblEstVivi']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblNroHabi']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblNroPers']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblServicios']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblbencre']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblcrevig']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblHabVivi']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblCodFami']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblPlaComp']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_lblMedPreg']",
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionVivienda = new List<string>()
                                    {
                                        "Tipo Vivienda", "Clase", "Especifique Vivienda", "Nombre Arrendador", "Teléfono Vivienda",
                                        "Perímetro VIvienda", "Estrato", "Número Habitaciones", "Número Personas Habitan", "Servicios",
                                        "Beneficiario Credito", "Crédito Vivienda Vigente", "Habitantes de la Vivienda", "Identificación",
                                        "Plan Complementario", "Plan Medicina Prepagada"
                                    };


                                    //Validación Emergentes campos
                                    for (int r = 0; r < 11; r++)
                                    {
                                        string campoName3 = selenium.CamposEmergentes(xpathVivienda[r], descripcionVivienda[r], file);
                                        Thread.Sleep(100);
                                        if (campoName3.Trim() != variablesEsperadas3[r].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas3[r] + " y el encontrado es: " + campoName3);
                                        }
                                        Thread.Sleep(100);
                                    }

                                    //SCROLL
                                    for (int k = 0; k < 4; k++)
                                    {
                                        Keyboard.SendKeys("{DOWN}");
                                    }

                                    //Validación Emergentes campos
                                    for (int r = 11; r < variablesEsperadas3.Count; r++)
                                    {
                                        string campoName3 = selenium.CamposEmergentes(xpathVivienda[r], descripcionVivienda[r], file);
                                        Thread.Sleep(100);
                                        if (campoName3.Trim() != variablesEsperadas3[r].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas3[r] + " y el encontrado es: " + campoName3);
                                        }
                                        Thread.Sleep(100);
                                    }

                                    Thread.Sleep(500);
                                    //SCROLL
                                    for (int k = 0; k < 5; k++)
                                    {
                                        Keyboard.SendKeys("{UP}");
                                    }

                                    /// -- TIEMPO LIBRE -- /// 

                                    //Clic en TiempoLibre
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_PnTieLib']/span");
                                    Thread.Sleep(100);

                                    //LISTA XPATH SERVICIO TIEMPO LIBRE
                                    List<string> xpathTiempoLibre = new List<string>() {
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_lblOfiDome']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_lblTieDest']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_lblOtrTide']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_lblRecReac']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_lblPerRecr']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_lblOtrPeri']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_lblDepOrte']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_lblPerDepo']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_lblOtrDepo']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_lblOtrTrab']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_lblPerTrab']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_lblOtrPetr']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionTiempoLibre = new List<string>()
                                    {
                                        "Oficios Dómesticos", "Tiempo Tarea", "Especifique Oficios Dómesticos", "Recreación", "Periocidad",
                                        "Especifique Recreación", "Deporte", "Periocidad Deporte", "Especifique Deporte", "Otro Trabajo", "Periocidad Otro Trabajo",
                                        "Especifique"
                                    };


                                    //Validación Emergentes campos
                                    for (int p = 0; p < variablesEsperadas4.Count; p++)
                                    {
                                        string campoName4 = selenium.CamposEmergentes(xpathTiempoLibre[p], descripcionTiempoLibre[p], file);
                                        Thread.Sleep(100);
                                        if (campoName4.Trim() != variablesEsperadas4[p].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas4[p] + " y el encontrado es: " + campoName4);
                                        }
                                        Thread.Sleep(100);
                                    }

                                    /// -- MEDIO DE DESPLAZAMIENTO -- ///  

                                    //Clic en Medio de Desplazamiento 
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnMedDesp']/span");
                                    Thread.Sleep(100);

                                    //LISTA XPATH SERVICIO MEDIO DE DESPLAZAMIENTO
                                    List<string> xpathMedioDesplazamiento = new List<string>() {
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_lblVehProp']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_lblVehPlac']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_Label2']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_lblTraPubl']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_lblNroBuse']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_lblOtrMetr']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_lblTraMuin']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_lblfretrl']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_Label4']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_lblTieDesp']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionMedioDesplazamientoe = new List<string>()
                                    {
                                        "Vehiculo Propio", "Placa Vehiculo", "Licencia Conducción", "Transporte Público", "Buses por Recorrido", "Especifique",
                                       "Transporte", "Traslados Jornada Laboral", "Medio Transporte", "Tiempo Diario Desplazamiento"
                                    };

                                    //Validación Emergentes campos
                                    for (int s = 0; s < variablesEsperadas5.Count; s++)
                                    {
                                        string campoName5 = selenium.CamposEmergentes(xpathMedioDesplazamiento[s], xpathMedioDesplazamiento[s], file);
                                        Thread.Sleep(100);
                                        if (campoName5.Trim() != variablesEsperadas5[s].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas5[s] + " y el encontrado es: " + campoName5);
                                        }
                                        Thread.Sleep(100);
                                    }

                                    /// -- OFICINA -- ///  

                                    //Clic en Medio de Oficina  
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnOfocina']/span");
                                    Thread.Sleep(100);

                                    //LISTA XPATH SERVICIO OFICINA
                                    List<string> xpathOficina = new List<string>() {
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_lblDirOfic']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_lblPaiOfic']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_lblDtoOfic']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_lblMpiOfic']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_lblTelOfic']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_lblExtOfic']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_lblTelFaxs']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcionOficina = new List<string>()
                                    {
                                       "Dirección Oficina", "Pais Oficina", "Departamento Oficina", "Municipio Oficina", "Teléfono Oficina",
                                       "Extenssión Oficina", "Fax Oficina"
                                    };

                                    //Validación Emergentes campos
                                    for (int d = 0; d < variablesEsperadas6.Count; d++)
                                    {
                                        string campoName6 = selenium.CamposEmergentes(xpathOficina[d], descripcionOficina[d], file);
                                        Thread.Sleep(100);
                                        if (campoName6.Trim() != variablesEsperadas6[d].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas6[d] + " y el encontrado es: " + campoName6);
                                        }
                                        Thread.Sleep(100);
                                    }

                                    /// -- DATOS COMPLEMENTARIOS -- ///   

                                    //Clic en Datos Complementarios   
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnDatCompl']/span");
                                    Thread.Sleep(100);

                                    //LISTA XPATH DATOS COMPLEMENTARIOS
                                    List<string> xpathDatosComplementarios = new List<string>() {
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnDatCompl_lblDependiente']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnDatCompl_lblParDisc']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnDatCompl_lblPertenece']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnDatCompl_lblGruSoci']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnDatCompl_lblPerReun']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnDatCompl_lblApoRgru']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnDatCompl_lblconesp']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnDatCompl_lblotroesp']",
                                        "//span[@id='ctl00_ContenidoPagina_TabContainer2_pnDatCompl_lblCOMLGTB']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> DatosComplementarios = new List<string>()
                                    {
                                       "Depennde de Personas Discapacitadas", "Parentesco", "Pertenece Grupo Social", "Cuál Grupo Social",
                                       "Periocidad de Reunion", "Aporte que Brinda al Grupo", "Condición Especial", "Cuál Condición Especial",
                                       "Pertenece a la comunidad LGBT"
                                    };

                                    //Validación Emergentes campos
                                    for (int f = 0; f < variablesEsperadas7.Count; f++)
                                    {
                                        string campoName7 = selenium.CamposEmergentes(xpathDatosComplementarios[f], DatosComplementarios[f], file);
                                        Thread.Sleep(100);
                                        if (campoName7.Trim() != variablesEsperadas7[f].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas7[f] + " y el encontrado es: " + campoName7);
                                        }
                                        Thread.Sleep(100);
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
        public void SmartPeople_frmBiSubFamLNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_6.SmartPeople_frmBiSubFamLNTC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                rows["ActualizarEsp"].ToString().Length != 0 && rows["ActualizarEsp"].ToString() != null &&
                                //Otros datos
                                rows["DatosBasicosFamiliaEsp"].ToString().Length != 0 && rows["DatosBasicosFamiliaEsp"].ToString() != null &&
                                rows["IdentificacionEsp"].ToString().Length != 0 && rows["IdentificacionEsp"].ToString() != null &&
                                rows["TipoDocumentoEspo"].ToString().Length != 0 && rows["TipoDocumentoEspo"].ToString() != null &&
                                rows["RegistroCivilEsp"].ToString().Length != 0 && rows["RegistroCivilEsp"].ToString() != null &&
                                rows["NombresEsp"].ToString().Length != 0 && rows["NombresEsp"].ToString() != null &&
                                rows["SegundoNombreEsp"].ToString().Length != 0 && rows["SegundoNombreEsp"].ToString() != null &&
                                rows["TipoRelacionEsp"].ToString().Length != 0 && rows["TipoRelacionEsp"].ToString() != null &&
                                rows["ApellidosEsp"].ToString().Length != 0 && rows["ApellidosEsp"].ToString() != null &&
                                rows["SegundoApellidoEsp"].ToString().Length != 0 && rows["SegundoApellidoEsp"].ToString() != null &&
                                rows["SexoEsp"].ToString().Length != 0 && rows["SexoEsp"].ToString() != null &&
                                rows["FechaNacimientoEsp"].ToString().Length != 0 && rows["FechaNacimientoEsp"].ToString() != null &&
                                rows["ViveEsp"].ToString().Length != 0 && rows["ViveEsp"].ToString() != null &&
                                rows["BeneficiarioEsp"].ToString().Length != 0 && rows["BeneficiarioEsp"].ToString() != null &&
                                rows["GrupoSangreEsp"].ToString().Length != 0 && rows["GrupoSangreEsp"].ToString() != null &&
                                rows["FactorSanguineoEsp"].ToString().Length != 0 && rows["FactorSanguineoEsp"].ToString() != null &&
                                rows["EstadoCivilEsp"].ToString().Length != 0 && rows["EstadoCivilEsp"].ToString() != null &&
                                rows["FechaMatrimonioEsp"].ToString().Length != 0 && rows["FechaMatrimonioEsp"].ToString() != null &&
                                rows["CiudadEsp"].ToString().Length != 0 && rows["CiudadEsp"].ToString() != null &&
                                rows["DireccionEsp"].ToString().Length != 0 && rows["DireccionEsp"].ToString() != null &&
                                rows["TelefonoEsp"].ToString().Length != 0 && rows["TelefonoEsp"].ToString() != null &&
                                rows["ActividadEsp"].ToString().Length != 0 && rows["ActividadEsp"].ToString() != null &&
                                rows["EstablecimientoEsp"].ToString().Length != 0 && rows["EstablecimientoEsp"].ToString() != null &&
                                rows["EscolaridadEsp"].ToString().Length != 0 && rows["EscolaridadEsp"].ToString() != null &&
                                rows["HobbiesEsp"].ToString().Length != 0 && rows["HobbiesEsp"].ToString() != null &&
                                rows["TrabajaOtraEntidadEsp"].ToString().Length != 0 && rows["TrabajaOtraEntidadEsp"].ToString() != null &&
                                rows["DiscapacidadEsp"].ToString().Length != 0 && rows["DiscapacidadEsp"].ToString() != null &&
                                rows["BeneficiarioCajaCompensacionEsp"].ToString().Length != 0 && rows["BeneficiarioCajaCompensacionEsp"].ToString() != null &&
                                rows["DependienteEmpleadoEsp"].ToString().Length != 0 && rows["DependienteEmpleadoEsp"].ToString() != null &&
                                rows["FechaRatificacionEsp"].ToString().Length != 0 && rows["FechaRatificacionEsp"].ToString() != null &&
                                rows["BeneficiarioEPSEsp"].ToString().Length != 0 && rows["BeneficiarioEPSEsp"].ToString() != null &&
                                rows["DocumentosEsp"].ToString().Length != 0 && rows["DocumentosEsp"].ToString() != null &&
                                rows["TipoDocumentoEsp"].ToString().Length != 0 && rows["TipoDocumentoEsp"].ToString() != null &&
                                //Url2
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                string ActualizarEsp = rows["ActualizarEsp"].ToString();
                                //Otros datos
                                string DatosBasicosFamiliaEsp = rows["DatosBasicosFamiliaEsp"].ToString();
                                string IdentificacionEsp = rows["IdentificacionEsp"].ToString();
                                string TipoDocumentoEspo = rows["TipoDocumentoEspo"].ToString();
                                string RegistroCivilEsp = rows["RegistroCivilEsp"].ToString();
                                string NombresEsp = rows["NombresEsp"].ToString();
                                string SegundoNombreEsp = rows["SegundoNombreEsp"].ToString();
                                string TipoRelacionEsp = rows["TipoRelacionEsp"].ToString();
                                string ApellidosEsp = rows["ApellidosEsp"].ToString();
                                string SegundoApellidoEsp = rows["SegundoApellidoEsp"].ToString();
                                string SexoEsp = rows["SexoEsp"].ToString();
                                string FechaNacimientoEsp = rows["FechaNacimientoEsp"].ToString();
                                string ViveEsp = rows["ViveEsp"].ToString();
                                string BeneficiarioEsp = rows["BeneficiarioEsp"].ToString();
                                string GrupoSangreEsp = rows["GrupoSangreEsp"].ToString();
                                string FactorSanguineoEsp = rows["FactorSanguineoEsp"].ToString();
                                string EstadoCivilEsp = rows["EstadoCivilEsp"].ToString();
                                string FechaMatrimonioEsp = rows["FechaMatrimonioEsp"].ToString();
                                string CiudadEsp = rows["CiudadEsp"].ToString();
                                string DireccionEsp = rows["DireccionEsp"].ToString();
                                string TelefonoEsp = rows["TelefonoEsp"].ToString();
                                string ActividadEsp = rows["ActividadEsp"].ToString();
                                string EstablecimientoEsp = rows["EstablecimientoEsp"].ToString();
                                string EscolaridadEsp = rows["EscolaridadEsp"].ToString();
                                string HobbiesEsp = rows["HobbiesEsp"].ToString();
                                string TrabajaOtraEntidadEsp = rows["TrabajaOtraEntidadEsp"].ToString();
                                string DiscapacidadEsp = rows["DiscapacidadEsp"].ToString();
                                string BeneficiarioCajaCompensacionEsp = rows["BeneficiarioCajaCompensacionEsp"].ToString();
                                string DependienteEmpleadoEsp = rows["DependienteEmpleadoEsp"].ToString();
                                string FechaRatificacionEsp = rows["FechaRatificacionEsp"].ToString();
                                string BeneficiarioEPSEsp = rows["BeneficiarioEPSEsp"].ToString();
                                string DocumentosEsp = rows["DocumentosEsp"].ToString();
                                string TipoDocumentoEsp = rows["TipoDocumentoEsp"].ToString();
                                //Url2
                                string url2 = rows["url2"].ToString();

                                List<string> variablesEsperadas = new List<string>() {  DatosBasicosFamiliaEsp, IdentificacionEsp, TipoDocumentoEspo, RegistroCivilEsp, NombresEsp, SegundoNombreEsp,
                                TipoRelacionEsp, ApellidosEsp, SegundoApellidoEsp, SexoEsp, FechaNacimientoEsp, ViveEsp, BeneficiarioEsp,
                                GrupoSangreEsp, FactorSanguineoEsp, EstadoCivilEsp, FechaMatrimonioEsp, CiudadEsp, DireccionEsp, TelefonoEsp,
                                ActividadEsp, EstablecimientoEsp, EscolaridadEsp, HobbiesEsp, TrabajaOtraEntidadEsp, DiscapacidadEsp,
                                BeneficiarioCajaCompensacionEsp, DependienteEmpleadoEsp, FechaRatificacionEsp, BeneficiarioEPSEsp,
                                DocumentosEsp, TipoDocumentoEsp };



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
                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://ophtsph:8085/".ToLower())
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

                                    ChromeDriver driver2 = selenium.returnDriver();

                                    driver2.Navigate().GoToUrl(url2);

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

                                    selenium.Screenshot("Titulos Descriptivos", true, file);

                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Actualizar
                                    string Actualizar = selenium.EmergenteBotones("btnGuardar");
                                    Thread.Sleep(100);
                                    if (Actualizar != ActualizarEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Actualizar es incorrecto, el esperado es: " + ActualizarEsp + " y el encontrado es: " + Actualizar);
                                    }
                                    Thread.Sleep(100);

                                    //////////  VALIDACION EMERGENTE DE CAMPOS  //////////

                                    //LISTA XPATH 
                                    List<string> xpath = new List<string>() {
                                        "//span[@id='ctl00_lblMenMarco']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodFami']",
                                        "//span[@id='ctl00_ContenidoPagina_lblTipIden']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodRcvl']",
                                        "//span[@id='ctl00_ContenidoPagina_lblNomFami1']",
                                        "//span[@id='ctl00_ContenidoPagina_lblNomFami2']",
                                        "//span[@id='ctl00_ContenidoPagina_lblTipRela']",
                                        "//span[@id='ctl00_ContenidoPagina_lblApeFami1']",
                                        "//span[@id='ctl00_ContenidoPagina_lblApeFami2']",
                                        "//span[@id='ctl00_ContenidoPagina_lblSexFami']",
                                        "//span[@id='ctl00_ContenidoPagina_kcfFecNaci_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_lblEstVida']",
                                        "//span[@id='ctl00_ContenidoPagina_lblPorBene']",
                                        "//span[@id='ctl00_ContenidoPagina_lblGruSang']",
                                        "//span[@id='ctl00_ContenidoPagina_lblFacSang']",
                                        "//span[@id='ctl00_ContenidoPagina_lblEstCivi']",
                                        "//span[@id='ctl00_ContenidoPagina_KcfFecMatr_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrDivip_lblDivPoli']",
                                        "//span[@id='ctl00_ContenidoPagina_lblDirFami']",
                                        "//span[@id='ctl00_ContenidoPagina_lblTelFami']",
                                        "//span[@id='ctl00_ContenidoPagina_lblTraEstu']",
                                        "//span[@id='ctl00_ContenidoPagina_lblSitEstu']",
                                        "//span[@id='ctl00_ContenidoPagina_lblGraEsco']",
                                        "//span[@id='ctl00_ContenidoPagina_lblHobFami']",
                                        "//span[@id='ctl00_ContenidoPagina_lblOtrEnti']",
                                        "//span[@id='ctl00_ContenidoPagina_lblEstDisc']",
                                        "//span[@id='ctl00_ContenidoPagina_lblBenCaco']",
                                        "//span[@id='ctl00_ContenidoPagina_lblFamDepe']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrlFecVncr_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_lblBenEps']",
                                        "//span[@id='ctl00_ContenidoPagina_lblAdjunto']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrTipoDocumento_lblTIP_DOCU']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcion = new List<string>()
                                    {
                                       "Datos Básicos Familiares","Identificación","Tipo Documento", "Registro Civil", "Nombres", "Segundo Nombre",
                                       "Tipo de Relación", "Apellidos", "Segundo Apellido", "Sexo", "Fecha de Nacimiento", "Vive", "Beneficiario",
                                       "Grupo de Sangre", "Factor Sanguíneo", "Estado Cívil", "Fecha Matrimonio/Convivencia", "Ciudad", "Dirección",
                                       "Teléfono", "Actividad", "Establecimiento", "Escolaridad", "Hobbies", "Trabaja con Otra Entidad", "Discapacidad",
                                       "Beneficiario a Caja de Compensación","Dependiente del Empleado", "Fecha Ratificación", "Beneficiario EPS", "Documentos",
                                       "Tipo de Documento"
                                    };

                                    //Validación Emergentes campos
                                    for (int i = 0; i < variablesEsperadas.Count; i++)
                                    {
                                        if (i != 16)
                                        {
                                            string campoName = selenium.CamposEmergentes(xpath[i], descripcion[i], file);
                                            Thread.Sleep(100);
                                            if (campoName.Trim() != variablesEsperadas[i].Trim())
                                            {
                                                errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas[i] + " y el encontrado es: " + campoName);
                                            }
                                            Thread.Sleep(100);

                                            if (i == 15)
                                            {
                                                for (int j = 0; j < 10; j++)
                                                {
                                                    Keyboard.SendKeys("{DOWN}");
                                                }
                                                Thread.Sleep(1000);
                                            }
                                        }


                                    }

                                    //SCROLL
                                    //selenium.Click("//div[@id='ctl00_ContenidoPagina_navgestionar']/div/div");
                                    for (int j = 0; j < 14; j++)
                                    {
                                        Keyboard.SendKeys("{UP}");
                                    }

                                    /////////// Validación TABS ////////
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

                                    if (elementList.Count > 0)
                                    {
                                        elementList[1].Click();
                                        Thread.Sleep(500);

                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
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
        public void SmartPeople_frmBpAtemeLNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_6.SmartPeople_frmBpAtemeLNTC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                rows["ActualizarEsp"].ToString().Length != 0 && rows["ActualizarEsp"].ToString() != null &&
                                //Otros datos
                                rows["DetalleSolicitudEsp"].ToString().Length != 0 && rows["DetalleSolicitudEsp"].ToString() != null &&
                                rows["NumContratoEsp"].ToString().Length != 0 && rows["NumContratoEsp"].ToString() != null &&
                                rows["IdentificacionEsp"].ToString().Length != 0 && rows["IdentificacionEsp"].ToString() != null &&
                                rows["NombreApellidosEsp"].ToString().Length != 0 && rows["NombreApellidosEsp"].ToString() != null &&
                                rows["ServicioEsp"].ToString().Length != 0 && rows["ServicioEsp"].ToString() != null &&
                                rows["FechaSolicitudEsp"].ToString().Length != 0 && rows["FechaSolicitudEsp"].ToString() != null &&
                                rows["PacienteEsp"].ToString().Length != 0 && rows["PacienteEsp"].ToString() != null &&
                                rows["DoctorEsp"].ToString().Length != 0 && rows["DoctorEsp"].ToString() != null &&
                                rows["FechaAsignacionEsp"].ToString().Length != 0 && rows["FechaAsignacionEsp"].ToString() != null &&
                                rows["HoraEsp"].ToString().Length != 0 && rows["HoraEsp"].ToString() != null &&
                                rows["EstadoCitaEsp"].ToString().Length != 0 && rows["EstadoCitaEsp"].ToString() != null &&
                                //Url2
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                string ActualizarEsp = rows["ActualizarEsp"].ToString();
                                //Otros datos
                                string DetalleSolicitudEsp = rows["DetalleSolicitudEsp"].ToString();
                                string NumContratoEsp = rows["NumContratoEsp"].ToString();
                                string IdentificacionEsp = rows["IdentificacionEsp"].ToString();
                                string NombreApellidosEsp = rows["NombreApellidosEsp"].ToString();
                                string ServicioEsp = rows["ServicioEsp"].ToString();
                                string FechaSolicitudEsp = rows["FechaSolicitudEsp"].ToString();
                                string PacienteEsp = rows["PacienteEsp"].ToString();
                                string DoctorEsp = rows["DoctorEsp"].ToString();
                                string FechaAsignacionEsp = rows["FechaAsignacionEsp"].ToString();
                                string HoraEsp = rows["HoraEsp"].ToString();
                                string EstadoCitaEsp = rows["EstadoCitaEsp"].ToString();
                                //Url2
                                string url2 = rows["url2"].ToString();

                                List<string> variablesEsperadas = new List<string>() {  DetalleSolicitudEsp, NumContratoEsp, IdentificacionEsp, NombreApellidosEsp, ServicioEsp,
                                FechaSolicitudEsp, PacienteEsp, DoctorEsp, FechaAsignacionEsp, HoraEsp, EstadoCitaEsp};

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
                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://ophtsph:8085/".ToLower())
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

                                    ChromeDriver driver2 = selenium.returnDriver();

                                    driver2.Navigate().GoToUrl(url2);

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

                                    selenium.Screenshot("Titulos Descriptivos", true, file);

                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Actualizar
                                    string Actualizar = selenium.EmergenteBotones("btnGuardar");
                                    Thread.Sleep(100);
                                    if (Actualizar != ActualizarEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Actualizar es incorrecto, el esperado es: " + ActualizarEsp + " y el encontrado es: " + Actualizar);
                                    }
                                    Thread.Sleep(100);

                                    //////////  VALIDACION EMERGENTE DE CAMPOS  //////////

                                    //LISTA XPATH 
                                    List<string> xpath = new List<string>() {
                                        "//span[@id='ctl00_lblMenMarco']",
                                        "//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "//span[@id='ctl00_ContenidoPagina_lblIndServ']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrlFecSoli_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_lblIndPers']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodDoct']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrlFecAsig_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrlHorAsig_lblHora']",
                                        "//span[@id='ctl00_ContenidoPagina_lblEstCita']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcions = new List<string>()
                                    {
                                       "Detalle de la Solicitud", "N. COntrato", "Identificación", "Nombres / Apellidos", "Servicio", "Fecha de Solicitud",
                                       "Paciente", "Doctor", "Fecha de Asignación", "Hora", "Estado de la Cita Medica"
                                    };

                                    //Validación Emergentes campos
                                    for (int i = 0; i < variablesEsperadas.Count; i++)
                                    {
                                        if (i != 3)
                                        {
                                            string campoName = selenium.CamposEmergentes(xpath[i], descripcions[i], file);
                                            Thread.Sleep(100);
                                            if (campoName.Trim() != variablesEsperadas[i].Trim())
                                            {
                                                errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas[i] + " y el encontrado es: " + campoName);
                                            }
                                            Thread.Sleep(100);
                                        }
                                    }

                                    /////////// Validación TABS ////////
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

                                    if (elementList.Count > 0)
                                    {
                                        elementList[1].Click();
                                        Thread.Sleep(500);

                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
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
        public void SmartPeople_frmDetReqpeHNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_6.SmartPeople_frmDetReqpeHNTC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                //rows["ActualizarEsp"].ToString().Length != 0 && rows["ActualizarEsp"].ToString() != null &&
                                //Otros datos
                                rows["AprobacionRequisicionesEsp"].ToString().Length != 0 && rows["AprobacionRequisicionesEsp"].ToString() != null &&
                                rows["CasoEsp"].ToString().Length != 0 && rows["CasoEsp"].ToString() != null &&
                                rows["AreaUsuariaEsp"].ToString().Length != 0 && rows["AreaUsuariaEsp"].ToString() != null &&
                                rows["FormaCoberturaEsp"].ToString().Length != 0 && rows["FormaCoberturaEsp"].ToString() != null &&
                                rows["FiltroSeleccionEsp"].ToString().Length != 0 && rows["FiltroSeleccionEsp"].ToString() != null &&
                                rows["ValidarVacanteEsp"].ToString().Length != 0 && rows["ValidarVacanteEsp"].ToString() != null &&
                                rows["FechaSolicitudEsp"].ToString().Length != 0 && rows["FechaSolicitudEsp"].ToString() != null &&
                                rows["CargoProveedorEsp"].ToString().Length != 0 && rows["CargoProveedorEsp"].ToString() != null &&
                                rows["NumPlazasEsp"].ToString().Length != 0 && rows["NumPlazasEsp"].ToString() != null &&
                                rows["FechaPosibleEsp"].ToString().Length != 0 && rows["FechaPosibleEsp"].ToString() != null &&
                                rows["CentroCostosEsp"].ToString().Length != 0 && rows["CentroCostosEsp"].ToString() != null &&
                                rows["CentroTrabajoEsp"].ToString().Length != 0 && rows["CentroTrabajoEsp"].ToString() != null &&
                                rows["HorarioTrabajoEsp"].ToString().Length != 0 && rows["HorarioTrabajoEsp"].ToString() != null &&
                                rows["JefeInmediatoEsp"].ToString().Length != 0 && rows["JefeInmediatoEsp"].ToString() != null &&
                                rows["DerechoDotacionEsp"].ToString().Length != 0 && rows["DerechoDotacionEsp"].ToString() != null &&
                                rows["PagoComisionesEsp"].ToString().Length != 0 && rows["PagoComisionesEsp"].ToString() != null &&
                                rows["PorcentajeComisionEsp"].ToString().Length != 0 && rows["PorcentajeComisionEsp"].ToString() != null &&
                                rows["PagoGarantizadoEsp"].ToString().Length != 0 && rows["PagoGarantizadoEsp"].ToString() != null &&
                                rows["PorcentajeGarantizadoEsp"].ToString().Length != 0 && rows["PorcentajeGarantizadoEsp"].ToString() != null &&
                                rows["TiempoGarantizado"].ToString().Length != 0 && rows["TiempoGarantizado"].ToString() != null &&
                                rows["OtrosPagosEsp"].ToString().Length != 0 && rows["OtrosPagosEsp"].ToString() != null &&
                                rows["ValorOtrosPagosEsp"].ToString().Length != 0 && rows["ValorOtrosPagosEsp"].ToString() != null &&
                                rows["TipoSalarioEsp"].ToString().Length != 0 && rows["TipoSalarioEsp"].ToString() != null &&
                                rows["SueldoBasicoEsp"].ToString().Length != 0 && rows["SueldoBasicoEsp"].ToString() != null &&
                                rows["HorasExtraEsp"].ToString().Length != 0 && rows["HorasExtraEsp"].ToString() != null &&
                                rows["MotivoSolicitudEsp"].ToString().Length != 0 && rows["MotivoSolicitudEsp"].ToString() != null &&
                                rows["TipoContratoEsp"].ToString().Length != 0 && rows["TipoContratoEsp"].ToString() != null &&
                                rows["FuncionarioReemplezarEsp"].ToString().Length != 0 && rows["FuncionarioReemplezarEsp"].ToString() != null &&
                                rows["CiudadEsp"].ToString().Length != 0 && rows["CiudadEsp"].ToString() != null &&
                                rows["ObservacionesSolicitudEsp"].ToString().Length != 0 && rows["ObservacionesSolicitudEsp"].ToString() != null &&
                                //Url2
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                //string ActualizarEsp = rows["ActualizarEsp"].ToString();
                                //Otros datos 
                                string AprobacionRequisicionesEsp = rows["AprobacionRequisicionesEsp"].ToString();
                                string CasoEsp = rows["CasoEsp"].ToString();
                                string AreaUsuariaEsp = rows["AreaUsuariaEsp"].ToString();
                                string FormaCoberturaEsp = rows["FormaCoberturaEsp"].ToString();
                                string FiltroSeleccionEsp = rows["FiltroSeleccionEsp"].ToString();
                                string ValidarVacanteEsp = rows["ValidarVacanteEsp"].ToString();
                                string FechaSolicitudEsp = rows["FechaSolicitudEsp"].ToString();
                                string CargoProveedorEsp = rows["CargoProveedorEsp"].ToString();
                                string NumPlazasEsp = rows["NumPlazasEsp"].ToString();
                                string FechaPosibleEsp = rows["FechaPosibleEsp"].ToString();
                                string CentroCostosEsp = rows["CentroCostosEsp"].ToString();
                                string CentroTrabajoEsp = rows["CentroTrabajoEsp"].ToString();
                                string HorarioTrabajoEsp = rows["HorarioTrabajoEsp"].ToString();
                                string JefeInmediatoEsp = rows["JefeInmediatoEsp"].ToString();
                                string DerechoDotacionEsp = rows["DerechoDotacionEsp"].ToString();
                                string PagoComisionesEsp = rows["PagoComisionesEsp"].ToString();
                                string PorcentajeComisionEsp = rows["PorcentajeComisionEsp"].ToString();
                                string PagoGarantizadoEsp = rows["PagoGarantizadoEsp"].ToString();
                                string PorcentajeGarantizadoEsp = rows["PorcentajeGarantizadoEsp"].ToString();
                                string TiempoGarantizado = rows["TiempoGarantizado"].ToString();
                                string OtrosPagosEsp = rows["OtrosPagosEsp"].ToString();
                                string ValorOtrosPagosEsp = rows["ValorOtrosPagosEsp"].ToString();
                                string TipoSalarioEsp = rows["TipoSalarioEsp"].ToString();
                                string SueldoBasicoEsp = rows["SueldoBasicoEsp"].ToString();
                                string HorasExtraEsp = rows["HorasExtraEsp"].ToString();
                                string MotivoSolicitudEsp = rows["MotivoSolicitudEsp"].ToString();
                                string TipoContratoEsp = rows["TipoContratoEsp"].ToString();
                                string FuncionarioReemplezarEsp = rows["FuncionarioReemplezarEsp"].ToString();
                                string CiudadEsp = rows["CiudadEsp"].ToString();
                                string ObservacionesSolicitudEsp = rows["ObservacionesSolicitudEsp"].ToString();

                                //Url2
                                string url2 = rows["url2"].ToString();

                                List<string> variablesEsperadas = new List<string>() {  AprobacionRequisicionesEsp, CasoEsp, AreaUsuariaEsp, FormaCoberturaEsp, FiltroSeleccionEsp, ValidarVacanteEsp,
                                    FechaSolicitudEsp, CargoProveedorEsp, NumPlazasEsp, FechaPosibleEsp, CentroCostosEsp, CentroTrabajoEsp, HorarioTrabajoEsp,
                                    JefeInmediatoEsp, DerechoDotacionEsp, PagoComisionesEsp, PorcentajeComisionEsp, PagoGarantizadoEsp, PorcentajeGarantizadoEsp,
                                    TiempoGarantizado, OtrosPagosEsp, ValorOtrosPagosEsp, TipoSalarioEsp, SueldoBasicoEsp, HorasExtraEsp, MotivoSolicitudEsp,
                                    TipoContratoEsp, FuncionarioReemplezarEsp, CiudadEsp, ObservacionesSolicitudEsp };


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
                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://ophtsph:8085/".ToLower())
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

                                    ChromeDriver driver2 = selenium.returnDriver();

                                    driver2.Navigate().GoToUrl(url2);

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

                                    selenium.Screenshot("Titulos Descriptivos", true, file);

                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);

                                    /*//Validación emergentes Boton Actualizar
                                    string Actualizar = selenium.EmergenteBotones("btnGuardar");
                                    Thread.Sleep(100);
                                    if (Actualizar != ActualizarEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Actualizar es incorrecto, el esperado es: " + ActualizarEsp + " y el encontrado es: " + Actualizar);
                                    }
                                    Thread.Sleep(100);*/

                                    //////////  VALIDACION EMERGENTE DE CAMPOS  //////////

                                    //LISTA XPATH 
                                    List<string> xpath = new List<string>() {
                                        "//span[@id='ctl00_lblMenMarco']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCasCont']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodGrse']",
                                        "//span[@id='ctl00_ContenidoPagina_lblForCobe']",
                                        "//span[@id='ctl00_ContenidoPagina_lblFilSele']",
                                        "//span[@id='ctl00_ContenidoPagina_lblValVaca']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrlFechaSoli_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodCarp']",
                                        "//span[@id='ctl00_ContenidoPagina_lblNroPlaz']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrlFechaPoin_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodCcos']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodCenp']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodTurn']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodFrep']",
                                        "//span[@id='ctl00_ContenidoPagina_lblDerDota']",
                                        "//span[@id='ctl00_ContenidoPagina_lblPagComi']",
                                        "//span[@id='ctl00_ContenidoPagina_lblPorComi']",
                                        "//span[@id='ctl00_ContenidoPagina_lblPagGara']",
                                        "//span[@id='ctl00_ContenidoPagina_lblPorGara']",
                                        "//span[@id='ctl00_ContenidoPagina_lblTieGara']",
                                        "//span[@id='ctl00_ContenidoPagina_lblOtrPago']",
                                        "//span[@id='ctl00_ContenidoPagina_lblValOtro']",
                                        "//span[@id='ctl00_ContenidoPagina_lblTipSala']",
                                        "//span[@id='ctl00_ContenidoPagina_lblSueProp']",
                                        "//span[@id='ctl00_ContenidoPagina_lblHraExtf']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodMoti']",
                                        "//span[@id='ctl00_ContenidoPagina_lblTipCont']",
                                        "//span[@id='ctl00_ContenidoPagina_lblFunReem']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrDivipEMPR_lblDivPoli']",
                                        "//span[@id='ctl00_ContenidoPagina_KtxtObserSoli_lblTexto']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcions = new List<string>()
                                    {
                                       "Aprobación de Requisiciones", "Caso", "Área Usuaria", "Forma de Cobertura Vacante", "FIltro de Selección",
                                       "Validar Vacantes", "Fecha Solicitud", "Cargo a Proveer", "Número Plazas", "Fecha Posible Ingreso", "Cemtro de Costo",
                                       "Centro de Trabajo", "Horario de Trabajo", "Jefe Inmediato", "Derecho a Dotación", "Pago a Comisiones", "Porcentaje Comisión",
                                       "Pago Garantizado", "Porcentaje Garantizado", "Tiempo Garantizado", "Otros Pagos", "Valor Otros Pagos", "Tipo Salario",
                                       "Sueldo Básico", "Horas Extra Fijas", "Motivo de Solicitud", "Tipo de Contrato", "Funcionario a Reemplazar", "Ciudad",
                                       "Observaciones de la Solicitud"
                                    };

                                    //Validación Emergentes campos
                                    for (int i = 0; i < variablesEsperadas.Count; i++)
                                    {

                                        string campoName = selenium.CamposEmergentes(xpath[i], descripcions[i], file);
                                        Thread.Sleep(100);
                                        if (campoName.Trim() != variablesEsperadas[i].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas[i] + " y el encontrado es: " + campoName);
                                        }
                                        Thread.Sleep(100);
                                        if (i == 13)
                                        {
                                            for (int j = 0; j < 12; j++)
                                            {
                                                Keyboard.SendKeys("{DOWN}");
                                            }
                                            Thread.Sleep(1000);
                                        }

                                    }
                                    //Debugger.Launch();
                                    //SCROLL
                                    for (int k = 0; k < 17; k++)
                                    {
                                        Keyboard.SendKeys("{UP}");
                                    }
                                    /*///////////// VALIDACION DE TABS/////////////////

                                     ChromeDriver driver = selenium.returnDriver();
                                     List<IWebElement> elementList = new List<IWebElement>();
                                     List<IWebElement> elementListPagina = new List<IWebElement>();
                                     List<IWebElement> elementListPagina2 = new List<IWebElement>();
                                     Thread.Sleep(800);
                                     elementList.AddRange(driver.FindElements(By.XPath("//a[@id='btnGuardar']")));
                                     Thread.Sleep(500);
                                     selenium.Screenshot("Campos Requeridos", true, file);

                                     Thread.Sleep(100);

                                     if (elementList.Count > 0)
                                     {
                                         elementList[0].Click();
                                         Thread.Sleep(500);
                                         selenium.Click("//a[@id='ctl00_btnCerrar']");
                                         Thread.Sleep(500);
                                         elementListPagina.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));
                                         elementListPagina2.AddRange(driver.FindElements(By.XPath("//*[contains(@id,'ctl00_ContenidoPagina_')]")));

                                         if (elementListPagina.Count > 0)
                                         {
                                             int cont = 0;
                                             foreach (IWebElement pageEle in elementListPagina)
                                             {
                                                 cont++;
                                                 Thread.Sleep(800);

                                                 if (pageEle.TagName == "select" || pageEle.TagName == "input")
                                                 {
                                                     if (pageEle.Displayed && pageEle.Enabled)
                                                     {
                                                         String id = pageEle.GetAttribute("id");
                                                         if (id == "ctl00_ContenidoPagina_txtNomCurs")
                                                         {
                                                             errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El control de ID " + id + " no debe permitir la opcion de tabular");
                                                         }

                                                         Keyboard.SendKeys("{TAB}");
                                                         selenium.Screenshot("TAB", true, file);

                                                         Thread.Sleep(100);

                                                     }
                                                     else
                                                     {
                                                         Keyboard.SendKeys("{TAB}");
                                                         selenium.Screenshot("TAB", true, file);

                                                         Thread.Sleep(100);
                                                     }
                                                 }
                                                 if (cont >= 30)
                                                 {
                                                     break;
                                                 }
                                             }
                                         }
                                         if (elementListPagina2.Count > 0)
                                         {
                                             foreach (IWebElement pageEle in elementListPagina2)
                                             {
                                                 if (pageEle.TagName == "span" && pageEle.Displayed && pageEle.GetAttribute("Class") == "rfv")
                                                 {
                                                     String campo = pageEle.GetAttribute("id");
                                                     errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El control de ID " + campo + " es requerido");
                                                 }
                                             }

                                         }
                                     }*/
                                    //////////// VALIDACION DE TABS/////////////////

                                    //SIN BOTON DE GUARDAR

                                    ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);

                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
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
        public void SmartPeople_frmGnCronoActiNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_6.SmartPeople_frmGnCronoActiNTC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                //rows["ActualizarEsp"].ToString().Length != 0 && rows["ActualizarEsp"].ToString() != null &&
                                //Otros datos
                                rows["CronogramaGeneralEsp"].ToString().Length != 0 && rows["CronogramaGeneralEsp"].ToString() != null &&
                                rows["EmpresaEsp"].ToString().Length != 0 && rows["EmpresaEsp"].ToString() != null &&
                                rows["FechaDesdeEsp"].ToString().Length != 0 && rows["FechaDesdeEsp"].ToString() != null &&
                                rows["FechaHastaEsp"].ToString().Length != 0 && rows["FechaHastaEsp"].ToString() != null &&
                                rows["ConsultarIdentificacionEsp"].ToString().Length != 0 && rows["ConsultarIdentificacionEsp"].ToString() != null &&
                                rows["ConsultarNombresEsp"].ToString().Length != 0 && rows["ConsultarNombresEsp"].ToString() != null &&
                                rows["ConsultarApellidosEsp"].ToString().Length != 0 && rows["ConsultarApellidosEsp"].ToString() != null &&
                                //Url2
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                //string ActualizarEsp = rows["ActualizarEsp"].ToString();
                                //Otros datos 
                                string CronogramaGeneralEsp = rows["CronogramaGeneralEsp"].ToString();
                                string EmpresaEsp = rows["EmpresaEsp"].ToString();
                                string FechaDesdeEsp = rows["FechaDesdeEsp"].ToString();
                                string FechaHastaEsp = rows["FechaHastaEsp"].ToString();
                                string ConsultarIdentificacionEsp = rows["ConsultarIdentificacionEsp"].ToString();
                                string ConsultarNombresEsp = rows["ConsultarNombresEsp"].ToString();
                                string ConsultarApellidosEsp = rows["ConsultarApellidosEsp"].ToString();
                                //Url2
                                string url2 = rows["url2"].ToString();

                                List<string> variablesEsperadas = new List<string>() { CronogramaGeneralEsp, EmpresaEsp, FechaDesdeEsp, FechaHastaEsp, ConsultarIdentificacionEsp, ConsultarNombresEsp, ConsultarApellidosEsp };

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
                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://ophtsph:8085/".ToLower())
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

                                    ChromeDriver driver2 = selenium.returnDriver();

                                    driver2.Navigate().GoToUrl(url2);
                                    Thread.Sleep(6000);
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

                                    selenium.Screenshot("Titulos Descriptivos", true, file);

                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);


                                    //////////  VALIDACION EMERGENTE DE CAMPOS  //////////

                                    //LISTA XPATH 
                                    List<string> xpath = new List<string>() {
                                        "//span[@id='ctl00_lblMenMarco']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodEmpr']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrlFecInic_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrlFecFina_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_lblConsuCedEmpl']",
                                        "//span[@id='ctl00_ContenidoPagina_lblConsuNomEmpl']",
                                        "//span[@id='ctl00_ContenidoPagina_lblConsuAplEmpl']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcions = new List<string>()
                                    {
                                       "Cronograma General", "Empresa", "Fecha Desde", "Fecha Hasta", "Consultar por Identificación",
                                       "Consultar por Nombres", "Consultar por Apellidos"
                                    };

                                    //Validación Emergentes campos
                                    for (int i = 0; i < variablesEsperadas.Count; i++)
                                    {

                                        string campoName = selenium.CamposEmergentes(xpath[i], descripcions[i], file);
                                        Thread.Sleep(100);
                                        if (campoName.Trim() != variablesEsperadas[i].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas[i] + " y el encontrado es: " + campoName);
                                        }
                                        Thread.Sleep(100);
                                        if (i == 3)
                                        {
                                            for (int j = 0; j < 12; j++)
                                            {
                                                Keyboard.SendKeys("{DOWN}");
                                            }
                                            Thread.Sleep(1000);
                                        }
                                    }

                                    for (int j = 0; j < 13; j++)
                                    {
                                        Keyboard.SendKeys("{UP}");
                                    }
                                    Thread.Sleep(1000);

                                    //////////// VALIDACION DE TABS/////////////////

                                    //SIN BOTON DE GUARDAR

                                    ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);

                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
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
        public void SmartPeople_frmNmAusenLNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_6.SmartPeople_frmNmAusenLNTC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                //rows["ActualizarEsp"].ToString().Length != 0 && rows["ActualizarEsp"].ToString() != null &&
                                //Otros datos
                                rows["IngresoGestionSolicitudesEsp"].ToString().Length != 0 && rows["IngresoGestionSolicitudesEsp"].ToString() != null &&
                                rows["FiltroEsp"].ToString().Length != 0 && rows["FiltroEsp"].ToString() != null &&
                                rows["EmpresaEsp"].ToString().Length != 0 && rows["EmpresaEsp"].ToString() != null &&
                                rows["IdentificacionEsp"].ToString().Length != 0 && rows["IdentificacionEsp"].ToString() != null &&
                                rows["NumContratoEsp"].ToString().Length != 0 && rows["NumContratoEsp"].ToString() != null &&
                                rows["ConceptoEsp"].ToString().Length != 0 && rows["ConceptoEsp"].ToString() != null &&
                                rows["TipoAusentismoEsp"].ToString().Length != 0 && rows["TipoAusentismoEsp"].ToString() != null &&
                                rows["AusenciaEsp"].ToString().Length != 0 && rows["AusenciaEsp"].ToString() != null &&
                                rows["ResolucionEsp"].ToString().Length != 0 && rows["ResolucionEsp"].ToString() != null &&
                                rows["FechaInicioEsp"].ToString().Length != 0 && rows["FechaInicioEsp"].ToString() != null &&
                                rows["FechaFinalEsp"].ToString().Length != 0 && rows["FechaFinalEsp"].ToString() != null &&
                                rows["ObservacionesEsp"].ToString().Length != 0 && rows["ObservacionesEsp"].ToString() != null &&
                                //Url2
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                //string ActualizarEsp = rows["ActualizarEsp"].ToString();
                                //Otros datos 
                                string IngresoGestionSolicitudesEsp = rows["IngresoGestionSolicitudesEsp"].ToString();
                                string FiltroEsp = rows["FiltroEsp"].ToString();
                                string EmpresaEsp = rows["EmpresaEsp"].ToString();
                                string IdentificacionEsp = rows["IdentificacionEsp"].ToString();
                                string NumContratoEsp = rows["NumContratoEsp"].ToString();
                                string ConceptoEsp = rows["ConceptoEsp"].ToString();
                                string TipoAusentismoEsp = rows["TipoAusentismoEsp"].ToString();
                                string AusenciaEsp = rows["AusenciaEsp"].ToString();
                                string ResolucionEsp = rows["ResolucionEsp"].ToString();
                                string FechaInicioEsp = rows["FechaInicioEsp"].ToString();
                                string FechaFinalEsp = rows["FechaFinalEsp"].ToString();
                                string ObservacionesEsp = rows["ObservacionesEsp"].ToString();
                                //Url2
                                string url2 = rows["url2"].ToString();

                                List<string> variablesEsperadas = new List<string>() {  IngresoGestionSolicitudesEsp, FiltroEsp, EmpresaEsp, IdentificacionEsp, NumContratoEsp, ConceptoEsp,
                                TipoAusentismoEsp, AusenciaEsp, ResolucionEsp, FechaInicioEsp, FechaFinalEsp, ObservacionesEsp };

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
                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://ophtsph:8085/".ToLower())
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

                                    ChromeDriver driver2 = selenium.returnDriver();

                                    driver2.Navigate().GoToUrl(url2);

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

                                    selenium.Screenshot("Titulos Descriptivos", true, file);

                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);


                                    //////////  VALIDACION EMERGENTE DE CAMPOS  //////////

                                    //LISTA XPATH 
                                    List<string> xpath = new List<string>() {
                                        "//div[@id='printable']/span",
                                        "//span[@id='ctl00_ContenidoPagina_Label1']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodEmpr']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodConc']",
                                        "//span[@id='ctl00_ContenidoPagina_lblTipAuse']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodMaus']",
                                        "//span[@id='ctl00_ContenidoPagina_lblNumReso']",
                                        "//span[@id='ctl00_ContenidoPagina_KcfFecDesd_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_KcfFecHast_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_lblObsErva']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcions = new List<string>()
                                    {
                                       "Ingreso y Gestión de Solicitudes de Ausentismos", "Filtro", "Empresa", "Identificación", "Número Contrato",
                                       "Concepto", "Tipo Ausentismo", "Ausencia", "Resolución", "Fecha Inicio", "Fecha Final", "Observaciones"
                                    };

                                    //Validación Emergentes campos
                                    for (int i = 0; i < variablesEsperadas.Count; i++)
                                    {

                                        string campoName = selenium.CamposEmergentes(xpath[i], descripcions[i], file);
                                        Thread.Sleep(100);
                                        if (campoName.Trim() != variablesEsperadas[i].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas[i] + " y el encontrado es: " + campoName);
                                        }
                                        Thread.Sleep(100);
                                    }


                                    //////////// VALIDACION DE TABS/////////////////

                                    //SIN BOTON DE GUARDAR

                                    ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);

                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
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
        public void SmartPeople_frmNmCaltuDNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_6.SmartPeople_frmNmCaltuDNTC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                //rows["ActualizarEsp"].ToString().Length != 0 && rows["ActualizarEsp"].ToString() != null &&
                                //Otros datos
                                rows["CalendarioTurnosEsp"].ToString().Length != 0 && rows["CalendarioTurnosEsp"].ToString() != null &&
                                rows["TurnoEsp"].ToString().Length != 0 && rows["TurnoEsp"].ToString() != null &&
                                rows["ProgFechaInicialEsp"].ToString().Length != 0 && rows["ProgFechaInicialEsp"].ToString() != null &&
                                rows["FechaFinalEsp"].ToString().Length != 0 && rows["FechaFinalEsp"].ToString() != null &&
                                rows["FechaPagoEsp"].ToString().Length != 0 && rows["FechaPagoEsp"].ToString() != null &&
                                rows["FiltroPorEsp"].ToString().Length != 0 && rows["FiltroPorEsp"].ToString() != null &&
                                rows["IdentificacionEsp"].ToString().Length != 0 && rows["IdentificacionEsp"].ToString() != null &&
                                rows["NombreEsp"].ToString().Length != 0 && rows["NombreEsp"].ToString() != null &&
                                rows["ApellidoEsp"].ToString().Length != 0 && rows["ApellidoEsp"].ToString() != null &&
                                rows["EmpleadosEsp"].ToString().Length != 0 && rows["EmpleadosEsp"].ToString() != null &&
                                //Url2
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                //string ActualizarEsp = rows["ActualizarEsp"].ToString();
                                //Otros datos 
                                string CalendarioTurnosEsp = rows["CalendarioTurnosEsp"].ToString();
                                string TurnoEsp = rows["TurnoEsp"].ToString();
                                string ProgFechaInicialEsp = rows["ProgFechaInicialEsp"].ToString();
                                string FechaFinalEsp = rows["FechaFinalEsp"].ToString();
                                string FechaPagoEsp = rows["FechaPagoEsp"].ToString();
                                string FiltroPorEsp = rows["FiltroPorEsp"].ToString();
                                string IdentificacionEsp = rows["IdentificacionEsp"].ToString();
                                string NombreEsp = rows["NombreEsp"].ToString();
                                string ApellidoEsp = rows["ApellidoEsp"].ToString();
                                string EmpleadosEsp = rows["EmpleadosEsp"].ToString();

                                //Url2
                                string url2 = rows["url2"].ToString();

                                List<string> variablesEsperadas = new List<string>() {  CalendarioTurnosEsp, TurnoEsp, ProgFechaInicialEsp, FechaFinalEsp, FechaPagoEsp, FiltroPorEsp,
                                IdentificacionEsp, NombreEsp, ApellidoEsp, EmpleadosEsp };

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
                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://ophtsph:8085/".ToLower())
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

                                    ChromeDriver driver2 = selenium.returnDriver();

                                    driver2.Navigate().GoToUrl(url2);

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

                                    selenium.Screenshot("Titulos Descriptivos", true, file);

                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);


                                    //////////  VALIDACION EMERGENTE DE CAMPOS  //////////

                                    //LISTA XPATH 
                                    List<string> xpath = new List<string>() {
                                        "//span[@id='ctl00_lblMenMarco']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodEmpr']",
                                        "//span[@id='ctl00_ContenidoPagina_KcfFecProgIni_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_KcfFecProgFin_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_KcfFecPago_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_Label1']",
                                        "//span[@id='ctl00_ContenidoPagina_lblFiltroId']",
                                        "//span[@id='ctl00_ContenidoPagina_lblFiltroNombre']",
                                        "//span[@id='ctl00_ContenidoPagina_lblFiltroApellido']",
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcions = new List<string>()
                                    {
                                       "Calendario De Turnos", "Empresas", "Programación Fecha Inicial", "Fecha Final", "Fecha Pago",
                                       "Filtro Por", "Identificación", "Nombre", "Apellido"
                                    };

                                    //Validación Emergentes campos
                                    for (int i = 0; i < descripcions.Count; i++)
                                    {

                                        string campoName = selenium.CamposEmergentes(xpath[i], descripcions[i], file);
                                        Thread.Sleep(100);
                                        if (campoName.Trim() != variablesEsperadas[i].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas[i] + " y el encontrado es: " + campoName);
                                        }
                                        Thread.Sleep(100);
                                    }


                                    //////////// VALIDACION DE TABS/////////////////
                                    selenium.Click("//select[@id='ctl00_ContenidoPagina_ddlTurno']");
                                    for (int i = 0; i < 12; i++)
                                    {
                                        Keyboard.SendKeys("{TAB}");
                                        selenium.Screenshot("TAB", true, file);

                                        Thread.Sleep(100);
                                    }

                                    /*//SIN BOTON DE GUARDAR

                                    ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);

                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
                                            selenium.Screenshot("TAB", true, file);

                                            Thread.Sleep(100);
                                        }
                                    }*/

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
        public void SmartPeople_frmNmRepVacalNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_6.SmartPeople_frmNmRepVacalNTC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                //rows["ActualizarEsp"].ToString().Length != 0 && rows["ActualizarEsp"].ToString() != null &&
                                //Otros datos
                                rows["Reporte"].ToString().Length != 0 && rows["Reporte"].ToString() != null &&
                                //Url2
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                //string ActualizarEsp = rows["ActualizarEsp"].ToString();
                                //Otros datos 
                                string Reporte = rows["Reporte"].ToString();
                                //Url2
                                string url2 = rows["url2"].ToString();

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
                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://ophtsph:8085/".ToLower())
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

                                    ChromeDriver driver2 = selenium.returnDriver();

                                    driver2.Navigate().GoToUrl(url2);

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

                                    selenium.Screenshot("Titulos Descriptivos", true, file);

                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);


                                    //////////  VALIDACION EMERGENTE DE CAMPOS  //////////

                                    string campoName = selenium.CamposEmergentes("//span[@id='ctl00_lblMenMarco']", "Reporte", file);
                                    Thread.Sleep(100);
                                    if (campoName.Trim() != Reporte.Trim())
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + Reporte + " y el encontrado es: " + campoName);
                                    }
                                    Thread.Sleep(100);


                                    //////////// VALIDACION DE TABS/////////////////

                                    //SIN BOTON DE GUARDAR
                                    ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Actions action = new Actions(driver);
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        action.MoveByOffset(5, 5).Perform();
                                        Thread.Sleep(500);
                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
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
        public void SmartPeople_frmNmSouvrANTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_6.SmartPeople_frmNmSouvrANTC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                //rows["ActualizarEsp"].ToString().Length != 0 && rows["ActualizarEsp"].ToString() != null &&
                                //Otros datos
                                rows["SubsidioUVREsp"].ToString().Length != 0 && rows["SubsidioUVREsp"].ToString() != null &&
                                rows["NumContratoEsp"].ToString().Length != 0 && rows["NumContratoEsp"].ToString() != null &&
                                rows["IdentificacionEsp"].ToString().Length != 0 && rows["IdentificacionEsp"].ToString() != null &&
                                rows["NombreApellidoEsp"].ToString().Length != 0 && rows["NombreApellidoEsp"].ToString() != null &&
                                rows["CodInternoEsp"].ToString().Length != 0 && rows["CodInternoEsp"].ToString() != null &&
                                rows["MontoMaxCreditoEsp"].ToString().Length != 0 && rows["MontoMaxCreditoEsp"].ToString() != null &&
                                rows["ValorRealCreditoEsp"].ToString().Length != 0 && rows["ValorRealCreditoEsp"].ToString() != null &&
                                rows["MontoCalculoCreditoEsp"].ToString().Length != 0 && rows["MontoCalculoCreditoEsp"].ToString() != null &&
                                rows["TasaInteresEsp"].ToString().Length != 0 && rows["TasaInteresEsp"].ToString() != null &&
                                rows["AñosPendientesEsp"].ToString().Length != 0 && rows["AñosPendientesEsp"].ToString() != null &&
                                rows["FechaSolicitudEsp"].ToString().Length != 0 && rows["FechaSolicitudEsp"].ToString() != null &&
                                rows["TipoSubsidioEsp"].ToString().Length != 0 && rows["TipoSubsidioEsp"].ToString() != null &&
                                rows["EntidadEsp"].ToString().Length != 0 && rows["EntidadEsp"].ToString() != null &&
                                //Url2
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                //string ActualizarEsp = rows["ActualizarEsp"].ToString();
                                //Otros datos 
                                string SubsidioUVREsp = rows["SubsidioUVREsp"].ToString();
                                string NumContratoEsp = rows["NumContratoEsp"].ToString();
                                string IdentificacionEsp = rows["IdentificacionEsp"].ToString();
                                string NombreApellidoEsp = rows["NombreApellidoEsp"].ToString();
                                string CodInternoEsp = rows["CodInternoEsp"].ToString();
                                string MontoMaxCreditoEsp = rows["MontoMaxCreditoEsp"].ToString();
                                string ValorRealCreditoEsp = rows["ValorRealCreditoEsp"].ToString();
                                string MontoCalculoCreditoEsp = rows["MontoCalculoCreditoEsp"].ToString();
                                string TasaInteresEsp = rows["TasaInteresEsp"].ToString();
                                string AñosPendientesEsp = rows["AñosPendientesEsp"].ToString();
                                string FechaSolicitudEsp = rows["FechaSolicitudEsp"].ToString();
                                string TipoSubsidioEsp = rows["TipoSubsidioEsp"].ToString();
                                string EntidadEsp = rows["EntidadEsp"].ToString();
                                //Url2
                                string url2 = rows["url2"].ToString();

                                List<string> variablesEsperadas = new List<string>() {  SubsidioUVREsp, NumContratoEsp, IdentificacionEsp, NombreApellidoEsp, CodInternoEsp, MontoMaxCreditoEsp,
                                ValorRealCreditoEsp, MontoCalculoCreditoEsp, TasaInteresEsp, AñosPendientesEsp, FechaSolicitudEsp, TipoSubsidioEsp, EntidadEsp};

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
                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://ophtsph:8085/".ToLower())
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

                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);

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

                                    selenium.Screenshot("Titulos Descriptivos", true, file);

                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);


                                    //////////  VALIDACION EMERGENTE DE CAMPOS  //////////

                                    //LISTA XPATH 
                                    List<string> xpath = new List<string>() {
                                        "//div[@id='printable']/span",
                                        "//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodInte']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCreditoMaximo']",
                                        "//span[@id='ctl00_ContenidoPagina_lblValorRealCredito']",
                                        "//span[@id='ctl00_ContenidoPagina_lblMontoCalculoCredito']",
                                        "//span[@id='ctl00_ContenidoPagina_lblTasinte']",
                                        "//span[@id='ctl00_ContenidoPagina_lblAnosPendientes']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrlFecSoli_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodEnti']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodSucu']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcions = new List<string>()
                                    {
                                       "Subsidio UVR", "Número Contrato", "Identificación", "Nombres y Apellidos", "Código Interno", "Monto Máximo Crédito",
                                       "Valor Real Crédito", "Monto Calculo Crédito", "Tasa de Interes", "Años Pendientes", "Fecha Solicitud", "TIpo Subsidio UVR",
                                       "Entidad"
                                    };

                                    //Validación Emergentes campos
                                    for (int i = 0; i < variablesEsperadas.Count; i++)
                                    {

                                        string campoName = selenium.CamposEmergentes(xpath[i], descripcions[i], file);
                                        Thread.Sleep(100);
                                        if (campoName.Trim() != variablesEsperadas[i].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas[i] + " y el encontrado es: " + campoName);
                                        }
                                        Thread.Sleep(100);
                                    }


                                    //////////// VALIDACION DE TABS/////////////////

                                    //SIN BOTON DE GUARDAR

                                    ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);

                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
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
        public void SmartPeople_frmNmSupagANTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_6.SmartPeople_frmNmSupagANTC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                //rows["ActualizarEsp"].ToString().Length != 0 && rows["ActualizarEsp"].ToString() != null &&
                                //Otros datos
                                rows["CambioCuentasEsp"].ToString().Length != 0 && rows["CambioCuentasEsp"].ToString() != null &&
                                rows["NumContratoEsp"].ToString().Length != 0 && rows["NumContratoEsp"].ToString() != null &&
                                rows["IdentificacionEsp"].ToString().Length != 0 && rows["IdentificacionEsp"].ToString() != null &&
                                rows["NombreApellidoEsp"].ToString().Length != 0 && rows["NombreApellidoEsp"].ToString() != null &&
                                rows["CodInternoEsp"].ToString().Length != 0 && rows["CodInternoEsp"].ToString() != null &&
                                rows["MesPagoEsp"].ToString().Length != 0 && rows["MesPagoEsp"].ToString() != null &&
                                rows["AñoPagoEsp"].ToString().Length != 0 && rows["AñoPagoEsp"].ToString() != null &&
                                rows["TasaEsp"].ToString().Length != 0 && rows["TasaEsp"].ToString() != null &&
                                rows["FechaPagoEsp"].ToString().Length != 0 && rows["FechaPagoEsp"].ToString() != null &&
                                rows["SaldoAnteriorEsp"].ToString().Length != 0 && rows["SaldoAnteriorEsp"].ToString() != null &&
                                rows["ValorCuotaPagadaEsp"].ToString().Length != 0 && rows["ValorCuotaPagadaEsp"].ToString() != null &&
                                rows["InteresesPagadosEsp"].ToString().Length != 0 && rows["InteresesPagadosEsp"].ToString() != null &&
                                rows["NuevoSaldoEsp"].ToString().Length != 0 && rows["NuevoSaldoEsp"].ToString() != null &&
                                //Url2
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                //string ActualizarEsp = rows["ActualizarEsp"].ToString();
                                //Otros datos 
                                string CambioCuentasEsp = rows["CambioCuentasEsp"].ToString();
                                string NumContratoEsp = rows["NumContratoEsp"].ToString();
                                string IdentificacionEsp = rows["IdentificacionEsp"].ToString();
                                string NombreApellidoEsp = rows["NombreApellidoEsp"].ToString();
                                string CodInternoEsp = rows["CodInternoEsp"].ToString();
                                string MesPagoEsp = rows["MesPagoEsp"].ToString();
                                string AñoPagoEsp = rows["AñoPagoEsp"].ToString();
                                string TasaEsp = rows["TasaEsp"].ToString();
                                string FechaPagoEsp = rows["FechaPagoEsp"].ToString();
                                string SaldoAnteriorEsp = rows["SaldoAnteriorEsp"].ToString();
                                string ValorCuotaPagadaEsp = rows["ValorCuotaPagadaEsp"].ToString();
                                string InteresesPagadosEsp = rows["InteresesPagadosEsp"].ToString();
                                string NuevoSaldoEsp = rows["NuevoSaldoEsp"].ToString();
                                //Url2
                                string url2 = rows["url2"].ToString();

                                List<string> variablesEsperadas = new List<string>() {  CambioCuentasEsp, NumContratoEsp, IdentificacionEsp, NombreApellidoEsp, CodInternoEsp, MesPagoEsp, AñoPagoEsp,
                                TasaEsp, FechaPagoEsp, SaldoAnteriorEsp, ValorCuotaPagadaEsp, InteresesPagadosEsp, NuevoSaldoEsp};

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
                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://ophtsph:8085/".ToLower())
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

                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);

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

                                    selenium.Screenshot("Titulos Descriptivos", true, file);

                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);


                                    //////////  VALIDACION EMERGENTE DE CAMPOS  //////////

                                    //LISTA XPATH 
                                    List<string> xpath = new List<string>() {
                                        "//span[@id='ctl00_lblMenMarco']",
                                        "//span[@id='ctl00_ContenidoPagina_lblNroCont']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodEmpl']",
                                        "//span[@id='ctl00_ContenidoPagina_lblNomEmpl']",
                                        "//span[@id='ctl00_ContenidoPagina_lblCodInte']",
                                        "//span[@id='ctl00_ContenidoPagina_lblMesPago']",
                                        "//span[@id='ctl00_ContenidoPagina_lblAnoPago']",
                                        "//span[@id='ctl00_ContenidoPagina_lblTasa']",
                                        "//span[@id='ctl00_ContenidoPagina_KCtrlFecSoli_lblFecha']",
                                        "//span[@id='ctl00_ContenidoPagina_lblSalAnte']",
                                        "//span[@id='ctl00_ContenidoPagina_lblValCuot']",
                                        "//span[@id='ctl00_ContenidoPagina_lblIntPaga']",
                                        "//span[@id='ctl00_ContenidoPagina_lblSalNuev']"
                                    };

                                    //DESCRIPCION DE SCREENSHOTS
                                    List<string> descripcions = new List<string>()
                                    {
                                       "Cambio de Cuentas", "Número Contrato", "Identificación", "Nombres y Apellidos", "Código Interno", "Mes Pago", "Año Pago",
                                       "Tasa %", "Fecha Pago", "Saldo Anterior", "Valor Cuota Pagada", "Intereses Pagados", "Nuevo Saldo"
                                    };

                                    //Validación Emergentes campos
                                    for (int i = 0; i < variablesEsperadas.Count; i++)
                                    {

                                        string campoName = selenium.CamposEmergentes(xpath[i], descripcions[i], file);
                                        Thread.Sleep(100);
                                        if (campoName.Trim() != variablesEsperadas[i].Trim())
                                        {
                                            errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + variablesEsperadas[i] + " y el encontrado es: " + campoName);
                                        }
                                        Thread.Sleep(100);
                                    }


                                    //////////// VALIDACION DE TABS/////////////////

                                    //SIN BOTON DE GUARDAR

                                    ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);

                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
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
        public void SmartPeople_frmNmSiReFuMNTC()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test.SmartPeople_NTC_6.SmartPeople_frmNmSiReFuMNTC")
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
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                //Variables
                                rows["TituloEsp"].ToString().Length != 0 && rows["TituloEsp"].ToString() != null &&
                                rows["SubtituloEsp"].ToString().Length != 0 && rows["SubtituloEsp"].ToString() != null &&
                                rows["HomeEsp"].ToString().Length != 0 && rows["HomeEsp"].ToString() != null &&
                                //rows["ActualizarEsp"].ToString().Length != 0 && rows["ActualizarEsp"].ToString() != null &&
                                //Otros datos
                                rows["IngresoSimuladorRetencionEsp"].ToString().Length != 0 && rows["IngresoSimuladorRetencionEsp"].ToString() != null &&
                                //Url2
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null
                                )
                            {
                                //LOGIN
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                //Variables
                                string TituloEsp = rows["TituloEsp"].ToString();
                                string SubtituloEsp = rows["SubtituloEsp"].ToString();
                                string HomeEsp = rows["HomeEsp"].ToString();
                                //string ActualizarEsp = rows["ActualizarEsp"].ToString();
                                //Otros datos 
                                string IngresoSimuladorRetencionEsp = rows["IngresoSimuladorRetencionEsp"].ToString();
                                //Url2
                                string url2 = rows["url2"].ToString();

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
                                    if (url.ToLower() == "http://ophtsph:8088/".ToLower())
                                    {

                                        database = "ORA";
                                    }
                                    else if (url.ToLower() == "http://ophtsph:8085/".ToLower())
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

                                    ChromeDriver driver2 = selenium.returnDriver();
                                    driver2.Navigate().GoToUrl(url2);

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
                                    //Debugger.Launch();
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

                                    selenium.Screenshot("Titulos Descriptivos", true, file);

                                    Thread.Sleep(100);

                                    //Validación emergentes Boton Home
                                    string Home = selenium.EmergenteBotones("ctl00_btnHome");
                                    Thread.Sleep(100);
                                    if (Home != HomeEsp)
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre del botón Home es incorrecto, el esperado es: " + HomeEsp + " y el encontrado es: " + Home);
                                    }
                                    Thread.Sleep(100);


                                    //////////  VALIDACION EMERGENTE DE CAMPOS  //////////

                                    string campoName = selenium.CamposEmergentes("//span[@id='ctl00_lblMenMarco']", "Ingreso a Simulador de Retención de la Fuente", file);
                                    Thread.Sleep(100);
                                    if (campoName.Trim() != IngresoSimuladorRetencionEsp.Trim())
                                    {
                                        errorMessagesMetodo.Add("::::::::::::::::::::::" + "MSG: El Nombre Emergente es incorrecto, el esperado es: " + IngresoSimuladorRetencionEsp + " y el encontrado es: " + campoName);
                                    }
                                    Thread.Sleep(100);

                                    //////////// VALIDACION DE TABS/////////////////

                                    //SIN BOTON DE GUARDAR

                                    ChromeDriver driver = selenium.returnDriver();
                                    List<IWebElement> elementList = new List<IWebElement>();
                                    Thread.Sleep(800);
                                    elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));

                                    if (elementList.Count > 0)
                                    {
                                        elementList[0].Click();
                                        Thread.Sleep(500);

                                        foreach (IWebElement pageEle in elementList)
                                        {
                                            Keyboard.SendKeys("{TAB}");
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
