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
    public class Modulo_BI : FuncionesVitales
    {

        string Modulo = "Modulo_BI";
        string app = "SmartPeople";

        APISelenium selenium = new APISelenium();
        APIFuncionesVitales fv = new APIFuncionesVitales();
        APIDatabase db = new APIDatabase();

        public Modulo_BI()
        {

        }

        [TestMethod]
        public void BI_DatosAdicionales()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_DatosAdicionales")
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
                                 rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["Estatura"].ToString().Length != 0 && rows["Estatura"].ToString() != null &&
                                rows["Peso"].ToString().Length != 0 && rows["Peso"].ToString() != null &&
                                rows["Raza"].ToString().Length != 0 && rows["Raza"].ToString() != null &&
                                rows["Telefono"].ToString().Length != 0 && rows["Telefono"].ToString() != null &&
                                rows["TelefonoCon"].ToString().Length != 0 && rows["TelefonoCon"].ToString() != null &&
                                rows["Seguro"].ToString().Length != 0 && rows["Seguro"].ToString() != null &&
                                rows["Enfermendades"].ToString().Length != 0 && rows["Enfermendades"].ToString() != null &&
                                rows["Medicamentos"].ToString().Length != 0 && rows["Medicamentos"].ToString() != null &&
                                rows["Alergias"].ToString().Length != 0 && rows["Alergias"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["Hobbies"].ToString().Length != 0 && rows["Hobbies"].ToString() != null &&
                                rows["Propositos"].ToString().Length != 0 && rows["Propositos"].ToString() != null &&
                                rows["Estatura2"].ToString().Length != 0 && rows["Estatura2"].ToString() != null &&
                                rows["Peso2"].ToString().Length != 0 && rows["Peso2"].ToString() != null &&
                                rows["Raza2"].ToString().Length != 0 && rows["Raza2"].ToString() != null &&
                                rows["Telefono2"].ToString().Length != 0 && rows["Telefono2"].ToString() != null &&
                                rows["TelefonoCon2"].ToString().Length != 0 && rows["TelefonoCon2"].ToString() != null &&
                                rows["Seguro2"].ToString().Length != 0 && rows["Seguro2"].ToString() != null &&
                                rows["Enfermendades2"].ToString().Length != 0 && rows["Enfermendades2"].ToString() != null &&
                                rows["Medicamentos2"].ToString().Length != 0 && rows["Medicamentos2"].ToString() != null &&
                                rows["Alergias2"].ToString().Length != 0 && rows["Alergias2"].ToString() != null &&
                                rows["Observacion2"].ToString().Length != 0 && rows["Observacion2"].ToString() != null &&
                                rows["Hobbies2"].ToString().Length != 0 && rows["Hobbies2"].ToString() != null &&
                                rows["Propositos2"].ToString().Length != 0 && rows["Propositos2"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string Estatura = rows["Estatura"].ToString();
                                string Peso = rows["Peso"].ToString();
                                string Raza = rows["Raza"].ToString();
                                string Telefono = rows["Telefono"].ToString();
                                string TelefonoCon = rows["TelefonoCon"].ToString();
                                string Seguro = rows["Seguro"].ToString();
                                string Enfermendades = rows["Enfermendades"].ToString();
                                string Medicamentos = rows["Medicamentos"].ToString();
                                string Alergias = rows["Alergias"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string Hobbies = rows["Hobbies"].ToString();
                                string Propositos = rows["Propositos"].ToString();
                                string Estatura2 = rows["Estatura2"].ToString();
                                string Peso2 = rows["Peso2"].ToString();
                                string Raza2 = rows["Raza2"].ToString();
                                string Telefono2 = rows["Telefono2"].ToString();
                                string TelefonoCon2 = rows["TelefonoCon2"].ToString();
                                string Seguro2 = rows["Seguro2"].ToString();
                                string Enfermendades2 = rows["Enfermendades2"].ToString();
                                string Medicamentos2 = rows["Medicamentos2"].ToString();
                                string Alergias2 = rows["Alergias2"].ToString();
                                string Observacion2 = rows["Observacion2"].ToString();
                                string Hobbies2 = rows["Hobbies2"].ToString();
                                string Restricciones = rows["Restricciones"].ToString();
                                string Entidad = rows["Entidad"].ToString();
                                string NumeroPersonas = rows["NumeroPersonas"].ToString();
                                string Recreacion = rows["Recreacion"].ToString();
                                string Placa = rows["Placa"].ToString();
                                string Direccion = rows["Direccion"].ToString();
                                string Extension = rows["Extension"].ToString();
                                string Comunidad = rows["Comunidad"].ToString();
                                string Condicion = rows["Condicion"].ToString();

                             
                                try
                                {
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //limpiar procesos
                                    Process[] processes = Process.GetProcessesByName("chromedriver");
                                    if (processes.Length > 0)
                                    {
                                        for (int i = 0; i < processes.Length; i++)
                                        {
                                            processes[i].Kill();
                                        }
                                    }

                                    //Inicio
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[contains(.,'Datos Adicionales')]");
                                    selenium.Click("//a[contains(.,'Datos Adicionales')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Adicionales", true, file);

                                    //---------------------------------------------------Pestaña Otros Datos--------------------------------------------------------------
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnOtrDa']/span");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Otros datos", true, file);
                                    //Observacion
                                    selenium.Clear("//textarea[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_txtObsErva_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_txtObsErva_txtTexto']", Observacion);//Observacion
                                    Thread.Sleep(2000);
                                    //Hobies
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_KCtrlTxtHOB_EMPL_txtTexto']", Hobbies);//Hobbies
                                    Thread.Sleep(2000);
                                    //Propositos
                                    selenium.Clear("//textarea[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_KCtrlTxtMIS_PROP_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_KCtrlTxtMIS_PROP_txtTexto']", Propositos);//Propositos
                                    //Lateralidad
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_TabContainer2_body']");
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_rdblateralidad_0']");
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_TabContainer2_pnOtrDa_rdblateralidad_0']");
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_TabContainer2_body']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Ingresados", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnOtrDa']/span");
                                    Thread.Sleep(2000);

                                    //-------------------------------------------------Pestaña Salud----------------------------------------------
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnSalud']/span");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Salud", true, file);
                                    Thread.Sleep(2000);
                                    //Estatura
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtEstEmpl']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtEstEmpl']", Estatura);//Estatura188
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Estatura", true, file);
                                    //Peso
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtEstPeso']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtEstPeso']", Peso);//Peso34
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Peso", true, file);
                                    //Raza
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtCodRaza']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtCodRaza']", Raza);//Raza
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Raza", true, file);
                                    //Telefono
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtTelMedi']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtTelMedi']", Telefono);//Telefono333333333
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Telefono", true, file);
                                    //Seguro
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtEmpSegu']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtEmpSegu']", Seguro);//Seguro
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Seguro", true, file);
                                    Thread.Sleep(1000);
                                    //TelefonoCon
                                    selenium.Click("//div[contains(@id,'ctl00_pBotones')]/div");
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtTelEmer']");
                                    Thread.Sleep(1000);
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtTelEmer']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_txtTelEmer']", TelefonoCon);//Telefono333333333
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Telefono", true, file);
                                    //Medicamentos
                                    selenium.Clear("//textarea[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_KCtrlTxtMedicamentos_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_KCtrlTxtMedicamentos_txtTexto']", Medicamentos);//Medicamentos
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Medicamentos", true, file);
                                    //Alergias
                                    selenium.Clear("//textarea[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_KCtrlTxtOtros_txtTexto']");
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_TabContainer2_pnSalud_KCtrlTxtOtros_txtTexto']", Alergias);//Alergias
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Alergias", true, file);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Datos Ingresados", true, file);
                                    selenium.Scroll("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnOtrDa']/span");

                                    //-----------------------------------Pestaña Servicios Salud----------------------------------------------
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnSerSalu']/span");
                                    Thread.Sleep(2000);
                                    //Grupo Apoyo
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_Checkfor']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Grupo Apoyo", true, file);
                                    //Restricciones Medicas
                                    selenium.Scroll("//textarea[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_KCtrlTxtRES_MEDS_txtTexto']");
                                    selenium.Clear("//textarea[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_KCtrlTxtRES_MEDS_txtTexto']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_KCtrlTxtRES_MEDS_txtTexto']", Restricciones);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Restricciones", true, file);
                                    //Entidad
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_txtENTMEDS']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnSerSalu_txtENTMEDS']", Entidad);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Entidad", true, file);
                                    selenium.Scroll("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnOtrDa']/span");
                                    //------------------------------------------------Pestaña Vivienda--------------------------------------------------------------------
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnVivienda']/span");
                                    Thread.Sleep(2000);
                                    //Vivienda Propia
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_rblVivProp_0']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Vivienda Propia", true, file);
                                    //Numero Personas Viven
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_txtNroPers']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_txtNroPers']", NumeroPersonas);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Numero Personas", true, file);
                                    //Servicios
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_txtNroPers']");
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_chkSerTele']");
                                    Thread.Sleep(5000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_txtNroPers']");
                                    selenium.Screenshot("TV", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnVivienda_chkSerTvsu']");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Telefono", true, file);
                                    selenium.Scroll("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnOtrDa']/span");
                                    //---------------------------------------------------Pestaña Tiempo Libre-------------------------------------------------
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_PnTieLib']/span");
                                    Thread.Sleep(2000);
                                    //Recreacion
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_txtRecReac']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_txtRecReac']", Recreacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Recreacion", true, file);
                                    //Periocidad Deporte
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_rblPerDepo_0']");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Periocidad Deporte", true, file);
                                    //Periocidad Otro Trabajo
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_TabContainer2_PnTieLib_rblPerTrab_0']");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Periocidad Otro Trabajo", true, file);
                                    selenium.Scroll("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnOtrDa']/span");
                                    //-----------------------------------------------Pestaña Medio desplazamiento---------------------------------------------------------
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnMedDesp']/span");
                                    Thread.Sleep(2000);
                                    //Placa Carro
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_txtVehPlac']");
                                    Thread.Sleep(2000);
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_txtVehPlac']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnMedDesp_txtVehPlac']", Placa);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Placa", true, file);
                                    selenium.Scroll("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnOtrDa']/span");
                                    //-----------------------------------------------Pestaña Oficina---------------------------------------------------------------------------
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnOfocina']");
                                    Thread.Sleep(2000);
                                    //Direccion Oficina
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_txtDirOfic']");
                                    Thread.Sleep(2000);
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_txtDirOfic']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_txtDirOfic']", Direccion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Direccion Oficina", true, file);
                                    //Extension Oficina
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_txtExtOfic']");
                                    Thread.Sleep(2000);
                                    selenium.Clear("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_txtExtOfic']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_TabContainer2_pnOfocina_txtExtOfic']", Extension);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Extension Oficina", true, file);
                                    selenium.Scroll("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnOtrDa']/span");
                                    //-----------------------------------------------Pestaña Datos Complementarios------------------------------------------------------------------
                                    selenium.Click("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnDatCompl']");
                                    Thread.Sleep(2000);
                                    //Condicion especial
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_TabContainer2_pnDatCompl_ddlConEspe']", Condicion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Condicion Especial", true, file);
                                    //Comunidad
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_TabContainer2_pnDatCompl_ddlCOMLGTB']", Comunidad);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Comunidad", true, file);
                                    selenium.Scroll("//a[@id='__tab_ctl00_ContenidoPagina_TabContainer2_pnOtrDa']/span");
                                    //----------------------------------------------------ACTUALIZAR DATOS--------------------------------------------------------------

                                    //Actualizar
                                    selenium.Click("//a[contains(@id,'btnActualizar')]");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Notificacion Actualizado Correcto", true, file);
                                    Thread.Sleep(5000);
                                    selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    Thread.Sleep(5000);
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
        public void BI_IngresoInformaciónEducaciónNoFormal()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_IngresoInformaciónEducaciónNoFormal")
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
                                 rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&

                                rows["Modalidad"].ToString().Length != 0 && rows["Modalidad"].ToString() != null &&
                                rows["NomEstudios"].ToString().Length != 0 && rows["NomEstudios"].ToString() != null &&
                                rows["NomEspecifico"].ToString().Length != 0 && rows["NomEspecifico"].ToString() != null &&
                                rows["Institucion"].ToString().Length != 0 && rows["Institucion"].ToString() != null &&
                                rows["FechaInicio"].ToString().Length != 0 && rows["FechaInicio"].ToString() != null &&
                                rows["FechaFin"].ToString().Length != 0 && rows["FechaFin"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Terminado"].ToString().Length != 0 && rows["Terminado"].ToString() != null &&
                                rows["TiempoEstudio"].ToString().Length != 0 && rows["TiempoEstudio"].ToString() != null &&
                                rows["Actualmente"].ToString().Length != 0 && rows["Actualmente"].ToString() != null &&
                                rows["UniTiem"].ToString().Length != 0 && rows["UniTiem"].ToString() != null &&
                                rows["EditNomEspecifico"].ToString().Length != 0 && rows["EditNomEspecifico"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["RmtEdnf"].ToString().Length != 0 && rows["RmtEdnf"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string user = rows["user"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string RmtEdnf = rows["RmtEdnf"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string Modalidad = rows["Modalidad"].ToString();
                                string NomEstudios = rows["NomEstudios"].ToString();
                                string NomEspecifico = rows["NomEspecifico"].ToString();
                                string Institucion = rows["Institucion"].ToString();
                                string FechaInicio = rows["FechaInicio"].ToString();
                                string FechaFin = rows["FechaFin"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Terminado = rows["Terminado"].ToString();
                                string TiempoEstudio = rows["TiempoEstudio"].ToString();
                                string Actualmente = rows["Actualmente"].ToString();
                                string UniTiem = rows["UniTiem"].ToString();
                                string EditNomEspecifico = rows["EditNomEspecifico"].ToString();
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
                                    //INICIO
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    string eliminarRegistro1 = $"delete BI_DTPNF where COD_EMPR='{CodEmpr}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRegistro1, database, user);
                                    string eliminarRegistro = $"delete bi_ednfo where cod_empl='{EmpleadoUser}' and act_usua='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                    //EDUCACION NO FORMAL
                                    Thread.Sleep(1000);
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Mi Educación No Formal')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mi Educación No Formal", true, file);
                                    //NUEVO
                                    bool existe = selenium.ExistControl("//a[@id='ctl00_btnNuevo']");
                                    if (existe)
                                    {
                                        selenium.Click("//a[@id='ctl00_btnNuevo']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Registro Nuevo", true, file);
                                    }
                                    //MODALIDAD
                                    selenium.SelectElementByName("//select[contains(@name,'ddlNomModi')]", Modalidad);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Modalidad", true, file);
                                    //NOMBRE ESTUDIOS
                                    selenium.SelectElementByName("//select[contains(@name,'ddlNomEstu')]", NomEstudios);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Estudios", true, file);
                                    //NOMBRE ESPECIFICO
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValMNomEspe_txtTexto']", NomEspecifico);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Nombre Específico", true, file);
                                    //INSTITUCION
                                    selenium.SendKeys("//input[contains(@name,'txtNomInst')]", Institucion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Institución", true, file);
                                    //FECHA INICIO
                                    selenium.Click("//input[contains(@id,'kcfFecInic')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'kcfFecInic')]", FechaInicio);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fecha Inicio", true, file);
                                    //FECHA FIN
                                    selenium.Click("//input[contains(@id,'kcfFecTerm')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'kcfFecTerm')]", FechaFin);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fecha Fin", true, file);
                                    //validaciones
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", ValPais);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    string Texto1 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    if (Texto1 != "")
                                    {
                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se eliminó el contenido del campo Ciudad al hacer TAB, el texto encontado es: " + Texto1);
                                    }

                                    //Validación 2 Departamento
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    string Texto2 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    if (Texto2 != "")
                                    {
                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se eliminó el contenido del campo Ciudad al hacer TAB, el texto encontado es: " + Texto2);
                                    }

                                    //Validación Caracteres especiales
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    string Texto3 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    if (Texto3 != "")
                                    {
                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se eliminó el contenido del campo Ciudad al hacer TAB, el texto encontado es: " + Texto3);
                                    }

                                    //Validación Exitosa
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    string Texto4 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    if (Texto4 == "")
                                    {
                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: Se eliminó el contenido del campo Ciudad al hacer TAB");
                                    }
                                    //TERMINDO ESTUDIOS
                                    selenium.Scroll("//select[contains(@name,'ddlEstTerm')]");
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@name,'ddlEstTerm')]", Terminado);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Terminado", true, file);
                                    //ESTUDIA ACTUALMENTE
                                    selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlEstActu']");
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlEstActu']", Actualmente);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Estudia Actualmente", true, file);
                                    //TIEMPO ESTUDIO
                                    selenium.Scroll("//input[contains(@name,'txtTieEstu')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@name,'txtTieEstu')]", TiempoEstudio);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tiempo estudio", true, file);
                                    //TIEMPO 
                                    selenium.SelectElementByName("//select[contains(@name,'ddlUniTiem')]", UniTiem);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tíempo", true, file);
                                    //selenium.returnDriver().ExecuteScript("arguments[0].scrollIntoView(true);", selenium.returnDriver().FindElement(By.XPath("//select[contains(@name,'ddlNomModi')]")));
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Datos Ingresados", true, file);
                                    //GUARDAR
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Registro Agregado", true, file);
                                    Thread.Sleep(500);
                                    //DETALLE
                                    selenium.Click("//*[@id='tblBiEdnfo']/tbody/tr/td[5]/a");
                                    Thread.Sleep(500);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Eliminar Registro", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(@id,'btnEliminar')]");
                                    Thread.Sleep(6000);
                                    selenium.Screenshot("Registro Eliminado", true, file);
                                    selenium.Close();
                                    Thread.Sleep(1000);
                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
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
        public void BI_Idiomas()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_Idiomas")
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
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&

                                rows["Idioma"].ToString().Length != 0 && rows["Idioma"].ToString() != null &&
                                rows["Habla"].ToString().Length != 0 && rows["Habla"].ToString() != null &&
                                rows["Lee"].ToString().Length != 0 && rows["Lee"].ToString() != null &&
                                rows["Escribe"].ToString().Length != 0 && rows["Escribe"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["Idioma2"].ToString().Length != 0 && rows["Idioma2"].ToString() != null &&

                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();

                                string Idioma = rows["Idioma"].ToString();
                                string Idioma2 = rows["Idioma2"].ToString();
                                string Habla = rows["Habla"].ToString();
                                string Lee = rows["Lee"].ToString();
                                string Escribe = rows["Escribe"].ToString();
                                string Observacion = rows["Observacion"].ToString();
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
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    IWebElement element = selenium.returnDriver().FindElement(By.XPath("//*[@id=\"MenuContex\"]/div[2]/div[1]/ul/li[1]/ul/li[9]/a"));
                                    IJavaScriptExecutor executor = (IJavaScriptExecutor)selenium.returnDriver();
                                    executor.ExecuteScript("arguments[0].click();", element);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Mis idiomas", true, file);

                                    int num = selenium.CountControl("//table[@id='ctl00_ContenidoPagina2_dgrBiEmidi']/tbody/tr[2]/td");
                                    if (num > 0)
                                    {
                                        string idiomaEnc = selenium.GetText("//table[@id='ctl00_ContenidoPagina2_dgrBiEmidi']/tbody/tr[2]/td");

                                        if (idiomaEnc != Idioma)
                                        {
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomIdio']", Idioma);
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlHabIdio']", Habla);
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlLeeIdio']", Lee);
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlEscIdio']", Escribe);
                                            selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtObsErva']", Observacion);
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Datos Ingresados", true, file);

                                            selenium.Click("//a[@id='btnGuardar']");
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Idioma Agregado", true, file);

                                            selenium.Click("//*[@id=\"ctl00_ContenidoPagina2_dgrBiEmidi_ctl02_LinkButton2\"]");
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Eliminar idioma", true, file);

                                        }
                                        else
                                        {
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomIdio']", Idioma2);
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlHabIdio']", Habla);
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlLeeIdio']", Lee);
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlEscIdio']", Escribe);
                                            selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtObsErva']", Observacion);
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Datos Ingresados", true, file);

                                            selenium.Click("//a[@id='btnGuardar']");
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Idioma Agregado", true, file);

                                            selenium.Click("//*[@id=\"ctl00_ContenidoPagina2_dgrBiEmidi_ctl03_LinkButton2\"]");
                                            Thread.Sleep(2000);
                                            selenium.Screenshot("Eliminar idioma", true, file);

                                        }
                                    }
                                    else
                                    {
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomIdio']", Idioma);
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlHabIdio']", Habla);
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlLeeIdio']", Lee);
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlEscIdio']", Escribe);
                                        selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtObsErva']", Observacion);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Datos Ingresados", true, file);

                                        selenium.Click("//a[@id='btnGuardar']");
                                        Thread.Sleep(6000);
                                        selenium.Screenshot("Idioma Agregado", true, file);

                                        selenium.Click("//*[@id=\"ctl00_ContenidoPagina2_dgrBiEmidi_ctl02_LinkButton2\"]");
                                        Thread.Sleep(4000);
                                        selenium.Screenshot("Eliminar idioma", true, file);

                                    }
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
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
        public void BI_DatosBásicos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_DatosBásicos")
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
                                 rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&

                                rows["Barrio"].ToString().Length != 0 && rows["Barrio"].ToString() != null &&
                                rows["NumCasa"].ToString().Length != 0 && rows["NumCasa"].ToString() != null &&
                                rows["Ruta"].ToString().Length != 0 && rows["Ruta"].ToString() != null &&
                                rows["Telefono"].ToString().Length != 0 && rows["Telefono"].ToString() != null &&
                                rows["Libreta"].ToString().Length != 0 && rows["Libreta"].ToString() != null &&

                                rows["Barrio2"].ToString().Length != 0 && rows["Barrio2"].ToString() != null &&
                                rows["NumCasa2"].ToString().Length != 0 && rows["NumCasa2"].ToString() != null &&
                                rows["Ruta2"].ToString().Length != 0 && rows["Ruta2"].ToString() != null &&
                                rows["Telefono2"].ToString().Length != 0 && rows["Telefono2"].ToString() != null &&
                                rows["Libreta2"].ToString().Length != 0 && rows["Libreta2"].ToString() != null &&

                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();

                                string Barrio = rows["Barrio"].ToString();
                                string NumCasa = rows["NumCasa"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string Telefono = rows["Telefono"].ToString();
                                string Libreta = rows["Libreta"].ToString();

                                string Barrio2 = rows["Barrio2"].ToString();
                                string NumCasa2 = rows["NumCasa2"].ToString();
                                string Ruta2 = rows["Ruta2"].ToString();
                                string Telefono2 = rows["Telefono2"].ToString();
                                string Libreta2 = rows["Libreta2"].ToString();

                                try
                                {
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //limpiar procesos
                                    Process[] processes = Process.GetProcessesByName("chromedriver");
                                    if (processes.Length > 0)
                                    {
                                        for (int i = 0; i < processes.Length; i++)
                                        {
                                            processes[i].Kill();
                                        }
                                    }
                                    //INICIO
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    //MIS DATOS BASICOS
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    selenium.Click("//a[contains(.,'Mis Datos Básicos')]");
                                    Thread.Sleep(1000);
                                    bool butClose = selenium.ExistControl("//a[@id='ctl00_btnCerrar']/i");
                                    if (butClose)
                                    {
                                        selenium.Click("//a[@id='ctl00_btnCerrar']/i");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Error", true, file);

                                    }

                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Datos Básicos", true, file);
                                    Thread.Sleep(2000);
                                    //BARRIO

                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtBarResiC']", Barrio);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Barrio", true, file);
                                    //NUMERO CASA

                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNumCasa']", NumCasa);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Número casa", true, file);
                                    //RUTA
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtRutResi']", Ruta);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Ruta", true, file);
                                    //TELEFONO

                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtTelResi']", Telefono);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Teléfono", true, file);
                                    SendKeys.SendWait("{ENTER}");
                                    selenium.ScrollTo("0", "200");
                                    Thread.Sleep(1000);
                                    if (database == "SQL")
                                    {
                                        //LIBRETA
                                        selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtNumLmil']");
                                        selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNumLmil']", Libreta);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Libreta", true, file);
                                    }
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Datos Agregados", true, file);

                                    //ACTUALIZAR
                                    Thread.Sleep(1000);
                                    selenium.Click("//a[@id='btnActualizar']");
                                    Thread.Sleep(10000);
                                    selenium.Screenshot("Datos Básicos Ingresados", true, file);
                                    //BOTON CERRAR
                                    Thread.Sleep(1000);
                                    selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    Thread.Sleep(3000);
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
        public void BI_IngresoMisFamiliares()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_IngresoMisFamiliares")
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
                                 rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&

                                rows["Id"].ToString().Length != 0 && rows["Id"].ToString() != null &&
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["SegNombre"].ToString().Length != 0 && rows["SegNombre"].ToString() != null &&
                                rows["Relacion"].ToString().Length != 0 && rows["Relacion"].ToString() != null &&
                                rows["Apellido"].ToString().Length != 0 && rows["Apellido"].ToString() != null &&
                                rows["SegApellido"].ToString().Length != 0 && rows["SegApellido"].ToString() != null &&
                                rows["Sexo"].ToString().Length != 0 && rows["Sexo"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["Beneficiario"].ToString().Length != 0 && rows["Beneficiario"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Sangre"].ToString().Length != 0 && rows["Sangre"].ToString() != null &&
                                rows["Sanguineo"].ToString().Length != 0 && rows["Sanguineo"].ToString() != null &&
                                rows["Ruta"].ToString().Length != 0 && rows["Ruta"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["TipoDocumAdjunto"].ToString().Length != 0 && rows["TipoDocumAdjunto"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string user = rows["user"].ToString();
                                string database = rows["database"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string Id = rows["Id"].ToString();
                                string Nombre = rows["Nombre"].ToString();
                                string SegNombre = rows["SegNombre"].ToString();
                                string Relacion = rows["Relacion"].ToString();
                                string Apellido = rows["Apellido"].ToString();
                                string SegApellido = rows["SegApellido"].ToString();
                                string Sexo = rows["Sexo"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string Beneficiario = rows["Beneficiario"].ToString();
                                string EstCivil = rows["EstCivil"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string Sangre = rows["Sangre"].ToString();
                                string Sanguineo = rows["Sanguineo"].ToString();
                                string TipoDocumAdjunto = rows["TipoDocumAdjunto"].ToString();
                                string Avenida = rows["Avenida"].ToString();
                                string numero = rows["numero"].ToString();
                                string letra = rows["letra"].ToString();
                                string bis = rows["bis"].ToString();
                                string letra2 = rows["letra2"].ToString();
                                string ubicacion = rows["ubicacion"].ToString();
                                string numero1 = rows["numero1"].ToString();
                                string letra3 = rows["letra3"].ToString();
                                string numero2 = rows["numero2"].ToString();
                                string ubicacion1 = rows["ubicacion1"].ToString();

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

                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    if (database == "ORA")
                                    {
                                        string eliminarRegistro = $"DELETE from bi_famil where cod_empl='{EmpleadoUser}' and cod_fami='{Id}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro1 = $"DELETE from bi_famil where cod_empl='{EmpleadoUser}' and cod_fami='123456789'";
                                        db.UpdateDeleteInsert(eliminarRegistro1, database, user);
                                    }
                                    else
                                    {
                                        string eliminarRegistro = $"DELETE from bi_famil where cod_empl='{EmpleadoUser}' and cod_fami='{Id}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro1 = $"DELETE from bi_famil where cod_empl='{EmpleadoUser}' and cod_fami='123456789'";
                                        db.UpdateDeleteInsert(eliminarRegistro1, database, user);
                                    }
                                    
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Mis Familiares')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//a[contains(.,'Mis Familiares')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Familiares", true, file);

                                    bool existe = selenium.ExistControl("//a[contains(@id,'ctl00_btnNuevo')]");
                                    if (existe)
                                    {
                                        selenium.Click("//a[contains(@id,'ctl00_btnNuevo')]");
                                    }
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nuevo", true, file);

                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_txtCodFami')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtCodFami')]", Id);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_txtNomFami1')]");
                                    Thread.Sleep(1500);
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtNomFami1')]", Nombre);
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_txtNomFami2')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtNomFami2')]", SegNombre);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlTipRela')]", Relacion);
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_txtApeFami1')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtApeFami1')]", Apellido);
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_txtApeFami2')]");
                                    Thread.Sleep(1500);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtApeFami2')]", SegApellido);
                                    Thread.Sleep(1500);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlSexFami')]", Sexo);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_kcfFecNaci_txtFecha')]", Fecha);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_txtPorBene')]", Beneficiario);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Datos Iniciales Familiar", true, file);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_ddlEstCivi')]", EstCivil);

                                 
                                    //Validación Exitosa
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivip_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivip_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Datos Ingresados", true, file);
                                    Thread.Sleep(1000);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    //CIUDAD
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KCtrDivip_txtDivPoli']", "COLOMBIA, CUNDINAMARCA, BOGOTA D.C.");
                                    Thread.Sleep(1500);
                                    selenium.Click("//*[@id='ctl00_pBotones']/div[1]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ciudad", true, file);
                                    //DIRECCION
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_lkbDirEcci']/span");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Direccion", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContentPopapModel_KCtrlDireccion1_Vias']", Avenida);
                                    selenium.SendKeys("//input[@id='inpnumero1']", numero);
                                    selenium.SelectElementByName("//select[@id='ddlletra1']", letra);
                                    selenium.SelectElementByName("//select[@id='ddlbis']", bis);
                                    selenium.SelectElementByName("//select[@id='ddlletra2']", letra2);
                                    selenium.SelectElementByName("//select[@id='ddlubi1']", ubicacion);
                                    selenium.SendKeys("//input[@id='inpnumero2']", numero1);
                                    selenium.SelectElementByName("//select[@id='ddlletra3']", letra3);
                                    selenium.SendKeys("//input[@id='inpnumero3']", numero2);
                                    selenium.SelectElementByName("//select[@id='ddlubi2']", ubicacion1);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Direccion Registrada", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContentPopapModel_btnterminar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Direccion Aplicada", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro Agregado", true, file);
                                    if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                    {
                                        selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    }
                                    //MI INFORMACION PERSONAL/FAMILIARES
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Mis Familiares')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Familiares", true, file);
                                    if (database == "SQL")
                                    {

                                        Thread.Sleep(500);
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_Familiares_ctl05_LinkButton1']/i");
                                    }
                                    else
                                    {
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_Familiares_ctl04_LinkButton1']/i");
                                    }

                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle", true, file);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Familiar a editar", true, file);
                                    Thread.Sleep(1500);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlEstCivi']", "Casado");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Estado Civil editado", true, file);
                                    selenium.Click("//a[contains(@id,'btnActualizar')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Registro Actualizado", true, file);
                                    Thread.Sleep(1000);
                                    Thread.Sleep(10000);
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
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

        public void BI_IngresoDocumentos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_IngresoDocumentos")
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
                                 rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&

                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["Numero"].ToString().Length != 0 && rows["Numero"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["FechaExpe"].ToString().Length != 0 && rows["FechaExpe"].ToString() != null &&
                                rows["FechaVen"].ToString().Length != 0 && rows["FechaVen"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["CodDocu"].ToString().Length != 0 && rows["CodDocu"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["TipoDocumAdjunto"].ToString().Length != 0 && rows["TipoDocumAdjunto"].ToString() != null &&
                                rows["Ruta"].ToString().Length != 0 && rows["Ruta"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string Descripcion = rows["Descripcion"].ToString();
                                string Numero = rows["Numero"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string FechaExpe = rows["FechaExpe"].ToString();
                                string FechaVen = rows["FechaVen"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string CodDocu = rows["CodDocu"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string TipoDocumAdjunto = rows["TipoDocumAdjunto"].ToString();
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
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //INICIO
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    //REGISTROS ELIMINAR
                                    string eliminarRegistro = $"DELETE bi_EMPDO where cod_empl='{EmpleadoUser}' and cod_docu='{CodDocu}' AND cod_empr='{CodEmpr}' AND act_usua='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRegistro, database, user);

                                    //MIS DOCUMENTOS
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Mis Documentos')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Mis Documentos')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Mis Documentos", true, file);
                                    //NUEVO
                                    bool existe = selenium.ExistControl("//a[@id='ctl00_btnNuevo']");
                                    if (existe)
                                    {
                                        selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    }

                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Agregar Nuevo", true, file);
                                    //IDENTIDAD
                                    Thread.Sleep(2000);
                                    selenium.Click("//*[@id=\"ctl00_ContenidoPagina_chbIndEntr\"]");
                                    Thread.Sleep(5000);
                                    //DESCRIPCION
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomDocu']", Descripcion);
                                    Thread.Sleep(2000);
                                    selenium.returnDriver().FindElement(By.Id("ctl00_ContenidoPagina_txtNumDocu")).Click();
                                    Thread.Sleep(2000);
                                    //NUMERO
                                    selenium.returnDriver().FindElement(By.Id("ctl00_ContenidoPagina_txtNumDocu")).SendKeys(Numero);
                                    Thread.Sleep(2000);
                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(2500);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.ScrollTo("0", "277");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Campo de Observación Requeridos", true, file);
                                    //CIUDAD
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipExpe_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Ciudad", true, file);
                                    //FECHAS
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_kcfFecExpe_txtFecha')]", FechaExpe);
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_kcfFecVenci_txtFecha')]", FechaVen);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Fechas", true, file);
                                    //OBSERVACIONES
                                    selenium.Scroll("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValObser_txtTexto']");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValObser_txtTexto']", Observacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Observaciones", true, file);
                                    //ADJUNTO
                                    for (int i = 0; i < 2; i++)
                                    {
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(1000);
                                    }
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Archivo adjunto", true, file);

                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro Agregado", true, file);

                                    if (database == "ORA")
                                    {
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_dtgBiEmpdo_ctl03_LinkButton1']/i");
                                    }
                                    else
                                    {
                                        selenium.Click("//a[@id='ctl00_ContenidoPagina_dtgBiEmpdo_ctl03_LinkButton1']/i");
                                    }
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Eliminar Registro", true, file);
                                    selenium.Click("//a[@id='ctl00_btnEliminar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Registro Eliminado", true, file);
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
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
        public void BI_TallasPorPrendaDeEmpleado()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_TallasPorPrendaDeEmpleado")
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
                                rows["Prenda4"].ToString().Length != 0 && rows["Prenda4"].ToString() != null &&
                                rows["Prenda3"].ToString().Length != 0 && rows["Prenda3"].ToString() != null &&
                                rows["Prenda2"].ToString().Length != 0 && rows["Prenda2"].ToString() != null &&
                                rows["Prenda1"].ToString().Length != 0 && rows["Prenda1"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null)
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string Prenda1 = rows["Prenda1"].ToString();
                                string Prenda2 = rows["Prenda2"].ToString();
                                string Prenda3 = rows["Prenda3"].ToString();
                                string Prenda4 = rows["Prenda4"].ToString();
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
                                    //INICIO
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    //ELIMINAR REGISTRO
                                    string eliminarRegistro = $"DELETE BI_EMTAL where COD_EMPL='{EmpleadoUser}' and cod_empr='{CodEmpr}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                    //TALLAS EMPLEADO
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'EF Tallas Empleado')]");
                                    selenium.Click("//a[contains(.,'EF Tallas Empleado')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Tallas Empleado", true, file);
                                    //PRENDA 1
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodPren']", Prenda1);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Prenda 1 seleccionada", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Prenda 1 Guardada", true, file);
                                    //ELIMINAR PRENDA 1
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrBiEmtal_ctl02_LinkButton1']/i");
                                    Thread.Sleep(1000);
                                    //PRENDA 2
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodPren']", Prenda2);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Prenda 2 seleccionada", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Prenda 2 Guardada", true, file);
                                    //PRENDA 2 ELIMINADA
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrBiEmtal_ctl02_LinkButton1']/i");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Prenda Eliminada", true, file);
                                    //PRENDA 3
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodPren']", Prenda3);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Prenda 3 seleccionada", true, file);
                                    //GUARDAR PRENDA 3
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Prenda 3 Guardada", true, file);
                                    //DETALLE
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_dgrBiEmtal_ctl02_lbSelDetalle']/i");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Detalle Prenda", true, file);
                                    //PRENDA
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodPren']", Prenda4);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Seleccionar Nueva prenda", true, file);
                                    //ACTUALIZAR
                                    selenium.Click("//a[@id='btnActualizar']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Notificacion prenda Actualizada", true, file);
                                    //NOTIFICACION
                                    selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
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
        public void BI_IngresoExperienciaLaboral()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_IngresoExperienciaLaboral")
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
                                rows["Empresa"].ToString().Length != 0 && rows["Empresa"].ToString() != null &&
                                rows["Direccion"].ToString().Length != 0 && rows["Direccion"].ToString() != null &&
                                rows["Telefono"].ToString().Length != 0 && rows["Telefono"].ToString() != null &&
                                rows["TipoEmpre"].ToString().Length != 0 && rows["TipoEmpre"].ToString() != null &&
                                rows["Correo"].ToString().Length != 0 && rows["Correo"].ToString() != null &&
                                rows["Salario"].ToString().Length != 0 && rows["Salario"].ToString() != null &&
                                rows["Empleo"].ToString().Length != 0 && rows["Empleo"].ToString() != null &&
                                rows["Cargo"].ToString().Length != 0 && rows["Cargo"].ToString() != null &&
                                rows["Dedicacion"].ToString().Length != 0 && rows["Dedicacion"].ToString() != null &&
                                rows["FechaIngre"].ToString().Length != 0 && rows["FechaIngre"].ToString() != null &&
                                rows["FechaReti"].ToString().Length != 0 && rows["FechaReti"].ToString() != null &&
                                rows["Personal"].ToString().Length != 0 && rows["Personal"].ToString() != null &&
                                rows["CargoDesemp"].ToString().Length != 0 && rows["CargoDesemp"].ToString() != null &&
                                rows["Area"].ToString().Length != 0 && rows["Area"].ToString() != null &&
                                rows["Contrato"].ToString().Length != 0 && rows["Contrato"].ToString() != null &&
                                rows["Retiro"].ToString().Length != 0 && rows["Retiro"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["JefeInmediato"].ToString().Length != 0 && rows["JefeInmediato"].ToString() != null &&
                                rows["CargoJefe"].ToString().Length != 0 && rows["CargoJefe"].ToString() != null &&
                                rows["EditDireccion"].ToString().Length != 0 && rows["EditDireccion"].ToString() != null &&
                                rows["Ruta"].ToString().Length != 0 && rows["Ruta"].ToString() != null &&
                                rows["RmtHvex"].ToString().Length != 0 && rows["RmtHvex"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null)
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string RmtHvex = rows["RmtHvex"].ToString();

                                string Empresa = rows["Empresa"].ToString();
                                string Direccion = rows["Direccion"].ToString();
                                string Telefono = rows["Telefono"].ToString();
                                string TipoEmpre = rows["TipoEmpre"].ToString();
                                string Correo = rows["Correo"].ToString();
                                string Salario = rows["Salario"].ToString();
                                string Empleo = rows["Empleo"].ToString();
                                string Cargo = rows["Cargo"].ToString();
                                string Dedicacion = rows["Dedicacion"].ToString();
                                string FechaIngre = rows["FechaIngre"].ToString();
                                string FechaReti = rows["FechaReti"].ToString();
                                string Personal = rows["Personal"].ToString();
                                string CargoDesemp = rows["CargoDesemp"].ToString();
                                string Area = rows["Area"].ToString();
                                string Contrato = rows["Contrato"].ToString();
                                string Retiro = rows["Retiro"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string JefeInmediato = rows["JefeInmediato"].ToString();
                                string CargoJefe = rows["Cargojefe"].ToString();
                                string EditDireccion = rows["EditDireccion"].ToString();
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
                                    //INICIO
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);
                                    //BORRA REGISTROS
                                    string eliminarRegistro1 = $"DELETE BI_ARHEX where cod_empr='{CodEmpr}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRegistro1, database, user);
                                    Thread.Sleep(2000);
                                    string eliminarRegistro2 = $"DELETE BI_DTPHV where cod_empr={CodEmpr} and ACT_USUA={EmpleadoUser}";
                                    db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                    string eliminarRegistro3 = $"DELETE bi_hvext where COD_EMPL={EmpleadoUser} and cod_empr={CodEmpr} and ACT_USUA={EmpleadoUser}";
                                    db.UpdateDeleteInsert(eliminarRegistro3, database, user);
                                    //MI EXPERIENCIA LABORAL
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(1500);
                                    selenium.Scroll("//a[contains(.,'Mi Experiencia Laboral')]");
                                    selenium.Click("//a[contains(.,'Mi Experiencia Laboral')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Mi Experiencia Laboral", true, file);
                                    //REGISTRO NUEVO
                                    bool existe = selenium.ExistControl("//a[@id='ctl00_btnNuevo']");
                                    if (existe)
                                    {
                                        selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    }
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Nuevo", true, file);
                                    //EMPRESA
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomEmpr']", Empresa);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empresa", true, file);
                                    //DIRECCION
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtDirEmpr']", Direccion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Dirección", true, file);
                                    //TELEFONO
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtTelEmpr']", Telefono);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Teléfono", true, file);
                                    //TIPO EMPRESA
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlTipEmpr']", TipoEmpre);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo Empresa", true, file);
                                    //CORREO
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtEntMail']", Correo);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Correo", true, file);
                                    //SALARIO
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtSalDemp']", Salario);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Salario", true, file);
                                    //EMPLEO
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlEmpActu']", Empleo);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Empleo", true, file);
                                    //CARGO
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCarEjec']", Cargo);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo", true, file);
                                    //DEDICACION
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDedIcac']", Dedicacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Dedicación", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_kcfFecIngr_txtFecha']");
                                    Thread.Sleep(500);
                                    //FECHA INGRESO
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_kcfFecIngr_txtFecha']", FechaIngre);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fecha Ingreso", true, file);
                                    //FECHA RETIRO
                                    selenium.Click("//input[@id='ctl00_ContenidoPagina_kcfFecReti_txtFecha']");
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_kcfFecReti_txtFecha']", FechaReti);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fecha Retiro", true, file);
                                    //MANEJA PERSONAL
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlManPers']", Personal);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Maneja Personal", true, file);
                                    //CARGO DESEMPEÑADO
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtCarDese']", CargoDesemp);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo Desempeñado", true, file);
                                    //AREA
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtDepEmpr']", Area);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Área", true, file);
                                    //CONTRATO
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlTipCont']", Contrato);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Contrato", true, file);
                                    //RETIRO
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlMotReti']", Retiro);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Retiro", true, file);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    //Validación 1 País
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]", ValPais);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    //Validación 2 Departamento
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    //Validación Caracteres especiales
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    //Validación Exitosa
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    //JEFE INMEDIATO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtJefInme']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtJefInme']", JefeInmediato);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Jefe Inmediato", true, file);
                                    //CARGO JEFE
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtCarJefe']", CargoJefe);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Cargo Jefe", true, file);
                                    //ADJUNTO
                                    selenium.Scroll("//*[@id=\"ctl00_ContenidoPagina_lblAdjunto\"]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//*[@id=\"ctl00_ContenidoPagina_lblAdjunto\"]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait(Ruta);
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Adjunto", true, file);
                                    selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Adjuntar Archivo", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']");
                                    Thread.Sleep(2000);
                                    selenium.ChangeAuxWindow();
                                    selenium.Click("//input[@id='ctl00_ContentPopapModel_btnNoProLog']");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Registro Exitoso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    //REGISTRO GUARDADO
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Mi Experiencia Laboral')]");
                                    selenium.Click("//a[contains(.,'Mi Experiencia Laboral')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Experiencia Registrada", true, file);
                                    //DETALLE
                                    selenium.Click("//*[@id='ctl00_ContenidoPagina_dgrBiHvext_ctl03_LinkButton1']/i");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Detalle", true, file);
                                    //ELIMINAR
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[@id='ctl00_btnEliminar']");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Registro Eliminado", true, file);
                                    selenium.Close();
         ////////////////////////////////////////////////////
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
        public void BI_IngresoInformaciónEducaciónFormal()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_IngresoInformaciónEducaciónFormal")
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
                                 rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["user"].ToString().Length != 0 && rows["user"].ToString() != null &&
                                rows["Modalidad"].ToString().Length != 0 && rows["Modalidad"].ToString() != null &&
                                rows["Estudios"].ToString().Length != 0 && rows["Estudios"].ToString() != null &&
                                rows["NomEspecifico"].ToString().Length != 0 && rows["NomEspecifico"].ToString() != null &&
                                rows["Institucion"].ToString().Length != 0 && rows["Institucion"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["Metodologia"].ToString().Length != 0 && rows["Metodologia"].ToString() != null &&
                                rows["FechaInicio"].ToString().Length != 0 && rows["FechaInicio"].ToString() != null &&
                                rows["FechaFin"].ToString().Length != 0 && rows["FechaFin"].ToString() != null &&
                                rows["TiempoEstudio"].ToString().Length != 0 && rows["TiempoEstudio"].ToString() != null &&
                                rows["TipoPeriodo"].ToString().Length != 0 && rows["TipoPeriodo"].ToString() != null &&
                                rows["Terminado"].ToString().Length != 0 && rows["Terminado"].ToString() != null &&
                                rows["Graduado"].ToString().Length != 0 && rows["Graduado"].ToString() != null &&
                                rows["FechaGrado"].ToString().Length != 0 && rows["FechaGrado"].ToString() != null &&
                                rows["Promedio"].ToString().Length != 0 && rows["Promedio"].ToString() != null &&
                                rows["Ruta"].ToString().Length != 0 && rows["Ruta"].ToString() != null &&
                                rows["RmtEdfo"].ToString().Length != 0 && rows["RmtEdfo"].ToString() != null &&
                                rows["CodEmpr"].ToString().Length != 0 && rows["CodEmpr"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["TipoDocumento"].ToString().Length != 0 && rows["TipoDocumento"].ToString() != null &&
                                rows["database"].ToString().Length != 0 && rows["database"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string url = rows["url"].ToString();
                                string database = rows["database"].ToString();
                                string user = rows["user"].ToString();
                                string CodEmpr = rows["CodEmpr"].ToString();
                                string RmtEdfo = rows["RmtEdfo"].ToString();

                                string Modalidad = rows["Modalidad"].ToString();
                                string Estudios = rows["Estudios"].ToString();
                                string NomEspecifico = rows["NomEspecifico"].ToString();
                                string Institucion = rows["Institucion"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string Metodologia = rows["Metodologia"].ToString();
                                string FechaInicio = rows["FechaInicio"].ToString();
                                string FechaFin = rows["FechaFin"].ToString();
                                string TiempoEstudio = rows["TiempoEstudio"].ToString();
                                string TipoPeriodo = rows["TipoPeriodo"].ToString();
                                string Terminado = rows["Terminado"].ToString();
                                string Graduado = rows["Graduado"].ToString();
                                string FechaGrado = rows["FechaGrado"].ToString();
                                string Promedio = rows["Promedio"].ToString();
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
                                    //INICIO
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2000);

                                    if (database == "SQL")
                                    {
                                        string eliminarRegistro = $"DELETE BI_DTPEF where cod_empr='{CodEmpr}' and act_usua='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro1 = $"DELETE bi_edfor where cod_empl='{EmpleadoUser}' and cod_empr='{CodEmpr}' and act_usua='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro1, database, user);
                                    }
                                    else
                                    {
                                        string eliminarRegistroOra = $"DELETE BI_DTPEF where cod_empr='{CodEmpr}' and act_usua='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistroOra, database, user);
                                        string eliminarRegistro1Ora = $"DELETE bi_edfor where cod_empl='{EmpleadoUser}' and cod_empr='{CodEmpr}' and act_usua='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro1Ora, database, user);
                                    }
                                    //EDUCACION FORMAL
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(500);
                                    selenium.Scroll("//a[contains(.,'Educacion Formal')]");
                                    Thread.Sleep(500);
                                    selenium.Click("//a[contains(.,'Educacion Formal')]");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Educación Formal", true, file);
                                    Thread.Sleep(2000);
                                    //NUEVO
                                    bool existe = selenium.ExistControl("//a[@id='ctl00_btnNuevo']");
                                    if (existe)
                                    {
                                        selenium.Click("//a[@id='ctl00_btnNuevo']");
                                    }
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Nuevo", true, file);
                                    //MODALIDAD
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomModi']", Modalidad);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Modalidad", true, file);
                                    //ESTUDIOS
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlNomEstu']", Estudios);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Estudios", true, file);
                                    //NOMBRE ESPECIFICO
                                    selenium.SendKeys("//textarea[@id='ctl00_ContenidoPagina_KCtrlTxtValMNomEspe_txtTexto']", NomEspecifico);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Nombre Específico", true, file);
                                    //INSTITUCION
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtNomInst']", Institucion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Estudios", true, file);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    //Validación 1 País
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", ValPais);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    string Texto1 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    if (Texto1 != "")
                                    {
                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se eliminó el contenido del campo Ciudad al hacer TAB, el texto encontado es: " + Texto1);
                                    }

                                    //Validación 2 Departamento
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    string Texto2 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    if (Texto2 != "")
                                    {
                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se eliminó el contenido del campo Ciudad al hacer TAB, el texto encontado es: " + Texto2);
                                    }

                                    //Validación Caracteres especiales
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    string Texto3 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    if (Texto3 != "")
                                    {
                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se eliminó el contenido del campo Ciudad al hacer TAB, el texto encontado es: " + Texto3);
                                    }

                                    //Validación Exitosa
                                    Thread.Sleep(500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    string Texto4 = selenium.GetTextFromTextBox("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    if (Texto4 == "")
                                    {
                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: Se eliminó el contenido del campo Ciudad al hacer TAB");
                                    }
                                    Thread.Sleep(500);

                                    if (database == "SQL")
                                    {
                                        selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlCodItem']", Metodologia);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Metodología", true, file);
                                    }

                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Datos Ingresados", true, file);
                                    
                                    //FECHA INICIO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_kcfFecInic_imgCalendario']/span");
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_kcfFecInic_imgCalendario']/span");
                                    Thread.Sleep(3000);
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_kcfFecInic_cetxtFecha_day_0_0']");
                                    Thread.Sleep(3000);
                                    //FECHA FIN
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_kcfFecTerm_imgCalendario']/span");
                                    Thread.Sleep(3000);
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_kcfFecTerm_cetxtFecha_day_0_0']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Fechas", true, file);
                                    //TERMINADO
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlEstTerm']", Terminado);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Terminado", true, file);
                                    //TIEMPO ESTUDIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtTieEstu']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtTieEstu']", TiempoEstudio);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tiempo Estudio", true, file);
                                    //TIPO PERIODO
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlUniTiem']", TipoPeriodo);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tipo Periodo", true, file);

                                    //PROMEDIO
                                    selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtProCarr']");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtProCarr']", Promedio);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Promedio", true, file);
                                    Thread.Sleep(2000);
                                    //CIUDAD
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    //ADJUNTO
                                    selenium.Scroll("//*[@id=\"ctl00_ContenidoPagina_lblAdjunto\"]");
                                    Thread.Sleep(1000);
                                    selenium.Click("//*[@id=\"ctl00_ContenidoPagina_lblAdjunto\"]");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(2000);
                                    SendKeys.SendWait(Ruta);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(1000);
                                    selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Adjuntar Archivo", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    bool exist = selenium.ExistControl("//a[contains(@id,'btnGuardar')]");
                                    if (exist)
                                    {
                                        selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    }

                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Registro Agregado", true, file);
                                    //DETALLE REGISTRO
                                    selenium.Click("//*[@id='tblBiEdfor']/tbody/tr/td[5]/a");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Detalle", true, file);
                                    //ELIMINAR REGISTRO
                                    selenium.Click("//a[contains(@id,'btnEliminar')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Eliminar Registro", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Close();

                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
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
        public void BI_ActualizarInformaciónDatosBásicos()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_ActualizarInformaciónDatosBásicos")
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
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["UdpDatoBasico"].ToString().Length != 0 && rows["UdpDatoBasico"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Modulo"].ToString().Length != 0 && rows["Modulo"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string UdpDatoBasico = rows["UdpDatoBasico"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina2 = rows["Maquina"].ToString();
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
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];

                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    List<string> errorMessagesMetodo = new List<string>();
                                    DateTime dateAndTime = DateTime.Now;
                                    string datetime = dateAndTime.ToString("ddMMyyyy_HHmmss");
                                    string UpdateData = UdpDatoBasico + datetime;

                                    //INICIO
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(2500);
                                    //DATOS BASICOS
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(2200);
                                    selenium.Click("//a[contains(.,'Mis Datos Básicos')]");
                                    Thread.Sleep(2500);
                                    selenium.Screenshot("Datos Básicos", true, file);
                                    //DATO A ACTUALIZAR
                                    selenium.Clear("//input[contains(@name,'txtBarResiC')]");
                                    selenium.SendKeys("//input[contains(@name,'txtBarResiC')]", UpdateData);
                                    Thread.Sleep(2500);
                                    selenium.Screenshot("Actualiza datos básicos", true, file);
                                    //ACTUALIZAR
                                    selenium.Click("//a[contains(@id,'btnActualizar')]");
                                    Thread.Sleep(2500);

                                    try
                                    {
                                        Thread.Sleep(3000);
                                        selenium.AcceptAlert();
                                    }
                                    catch
                                    {
                                    }
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Mensaje actualización de datos básicos", true, file);

                                    string Result = selenium.GetText("/html/body/div[1]/div/div[3]");

                                    if (Result != "Se actualizaron los datos correctamente.")
                                    {
                                        //errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: El resultado de la aplicacion no es el esperado...   Error: " + Result);
                                    }
                                    else
                                    {
                                        selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                        Thread.Sleep(1500);
                                        selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                        selenium.Click("//a[contains(.,'Mis Datos Básicos')]");
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Dato almacenado", true, file);

                                        string UpdateTextBox = selenium.GetTextFromTextBox("//input[contains(@name,'txtBarResiC')]");
                                        if (UpdateTextBox != UpdateData)
                                        {
                                            errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: El dato almacenado no corresponde: Almacenar: " + UpdateData + " Almacenado: " + UpdateTextBox);
                                        }
                                    }
                                    selenium.Close();
                                    Thread.Sleep(1500);
                                    fv.ConvertWordToPDF(file, database);  //LimpiarProcesos();
                                    ////////////////////////////////////////////////////
                                    ///
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
        public void BI_EducaciónFormal()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_EducaciónFormal")
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
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["Modalidad"].ToString().Length != 0 && rows["Modalidad"].ToString() != null &&
                                rows["NomProfesion"].ToString().Length != 0 && rows["NomProfesion"].ToString() != null &&
                                rows["NomEspecifico"].ToString().Length != 0 && rows["NomEspecifico"].ToString() != null &&
                                rows["NomInstitucion"].ToString().Length != 0 && rows["NomInstitucion"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["AnoInicio"].ToString().Length != 0 && rows["AnoInicio"].ToString() != null &&
                                rows["MesInicio"].ToString().Length != 0 && rows["MesInicio"].ToString() != null &&
                                rows["DiaInicio"].ToString().Length != 0 && rows["DiaInicio"].ToString() != null &&
                                rows["AnoFin"].ToString().Length != 0 && rows["AnoFin"].ToString() != null &&
                                rows["MesFin"].ToString().Length != 0 && rows["MesFin"].ToString() != null &&
                                rows["DiaFin"].ToString().Length != 0 && rows["DiaFin"].ToString() != null &&
                                rows["TiempoEducacion"].ToString().Length != 0 && rows["TiempoEducacion"].ToString() != null &&
                                rows["UniTiempo"].ToString().Length != 0 && rows["UniTiempo"].ToString() != null &&
                                rows["Terminado"].ToString().Length != 0 && rows["Terminado"].ToString() != null &&
                                rows["Graduado"].ToString().Length != 0 && rows["Graduado"].ToString() != null &&
                                rows["AnoGrado"].ToString().Length != 0 && rows["AnoGrado"].ToString() != null &&
                                rows["MesGrado"].ToString().Length != 0 && rows["MesGrado"].ToString() != null &&
                                rows["DiaGrado"].ToString().Length != 0 && rows["DiaGrado"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string Modalidad = rows["Modalidad"].ToString();
                                string NomProfesion = rows["NomProfesion"].ToString();
                                string NomEspecifico = rows["NomEspecifico"].ToString();
                                string NomInstitucion = rows["NomInstitucion"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string AnoInicio = rows["AnoInicio"].ToString();
                                string MesInicio = rows["MesInicio"].ToString();
                                string DiaInicio = rows["DiaInicio"].ToString();
                                string AnoFin = rows["AnoFin"].ToString();
                                string MesFin = rows["MesFin"].ToString();
                                string DiaFin = rows["DiaFin"].ToString();
                                string TiempoEducacion = rows["TiempoEducacion"].ToString();
                                string UniTiempo = rows["UniTiempo"].ToString();
                                string Terminado = rows["Terminado"].ToString();
                                string Graduado = rows["Graduado"].ToString();
                                string AnoGrado = rows["AnoGrado"].ToString();
                                string MesGrado = rows["MesGrado"].ToString();
                                string DiaGrado = rows["DiaGrado"].ToString();
                                string url = rows["url"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                try
                                {
                                    string user = "";
                                    string database = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //limpiar procesos
                                    Process[] processes = Process.GetProcessesByName("chromedriver");
                                    if (processes.Length > 0)
                                    {
                                        for (int i = 0; i < processes.Length; i++)
                                        {
                                            processes[i].Kill();
                                        }
                                    }
                                    //PARAMETRIZACION

                                    if (database == "SQL")
                                    {
                                        string eliminarRegistro = $"DELETE BI_DTPEF where cod_empr='9' and act_usua='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                        string eliminarRegistro1 = $"DELETE bi_edfor where cod_empl='{EmpleadoUser}' and cod_empr='9' and act_usua='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro1, database, user);
                                    }
                                    else
                                    {
                                        string eliminarRegistroOra = $"DELETE BI_DTPEF where cod_empr='421' and act_usua='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistroOra, database, user);
                                        string eliminarRegistro1Ora = $"DELETE bi_edfor where cod_empl='{EmpleadoUser}' and cod_empr='421' and act_usua='{EmpleadoUser}'";
                                        db.UpdateDeleteInsert(eliminarRegistro1Ora, database, user);
                                    }
                                    //INICIO
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Educacion Formal')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Educación formal", true, file);
                                    //MODALIDAD
                                    selenium.Scroll("//select[contains(@name,'Modi')]");
                                    selenium.SelectElementByName("//select[contains(@name,'Modi')]", Modalidad);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Modalidad", true, file);
                                    //PROFESION
                                    selenium.Scroll("//select[contains(@name,'NomEstu')]");
                                    selenium.SelectElementByName("//select[contains(@name,'NomEstu')]", NomProfesion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Profesión", true, file);
                                    //NOMBRE ESPECIFICO
                                    selenium.Scroll("//textarea[contains(@id,'KCtrlTxtValMNomEspe_txtTexto')]");
                                    selenium.SendKeys("//textarea[contains(@id,'KCtrlTxtValMNomEspe_txtTexto')]", NomEspecifico);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Nombre Específico", true, file);
                                    //NOMBRE INSTITUCION
                                    selenium.Scroll("//input[contains(@name,'NomInst')]");
                                    selenium.SendKeys("//input[contains(@name,'NomInst')]", NomInstitucion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Nombre Institución", true, file);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");



                                    //Validación Exitosa
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    

                                    //FECHA INICIO
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_kcfFecInic_imgCalendario']/span");
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_kcfFecInic_imgCalendario']/span");
                                    Thread.Sleep(3000);
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_kcfFecInic_cetxtFecha_day_0_0']");
                                    Thread.Sleep(3000);
                                    //FECHA FIN
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_kcfFecTerm_imgCalendario']/span");
                                    Thread.Sleep(3000);
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_kcfFecTerm_cetxtFecha_day_0_0']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Fechas", true, file);
                                    //TIEMPO EDUCACION
                                    selenium.Scroll("//select[contains(@name,'ddlEstTerm')]");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[contains(@name,'txtTieEstu')]", TiempoEducacion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tiempo educación", true, file);
                                    //TIEMPO
                                    selenium.Scroll("//select[contains(@name,'ddlUniTiem')]");
                                    selenium.SelectElementByName("//select[contains(@name,'ddlUniTiem')]", UniTiempo);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tiempo", true, file);
                                    //TERMINADO
                                    selenium.Scroll("//select[contains(@name,'ddlEstTerm')]");
                                    selenium.SelectElementByName("//select[contains(@name,'ddlEstTerm')]", Terminado);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Terminado", true, file);
                                    //GRADUADO
                                    selenium.Scroll("//select[contains(@name,'ddlGraDuad')]");
                                    selenium.SelectElementByName("//select[contains(@name,'ddlGraDuad')]", Graduado);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Graduado", true, file);
                                   
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(6000);


                                    if (selenium.ExistControl("//*[@id='tblBiEdfor']/tbody/tr/td[5]/a"))
                                    {
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Educación Formal Registrada", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Click("//*[@id='tblBiEdfor']/tbody/tr/td[5]/a");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Educación Formal a modificar", true, file);
                                       
                                        //CIUDAD
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                        Thread.Sleep(2000);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", Ciudad);
                                        Thread.Sleep(500);
                                        selenium.Click("//div[@id='ctl00_pBotones']/div");
                                        Thread.Sleep(2000);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", Ciudad);
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Ciudad", true, file);
                                         //PROMEDIO
                                        selenium.Scroll("//input[@id='ctl00_ContenidoPagina_txtProCarr']");
                                        Thread.Sleep(1000);
                                        selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_txtProCarr']", "3");
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Promedio", true, file);
                                        Thread.Sleep(2000);
                                        SendKeys.SendWait("{TAB}");
                                        Thread.Sleep(5000);
                                        ////ADJUNTO
                                        selenium.Scroll("//*[@id=\"ctl00_ContenidoPagina_lblAdjunto\"]");
                                        Thread.Sleep(2000);
                                        selenium.Click("//*[@id=\"ctl00_ContenidoPagina_lblAdjunto\"]");
                                        Thread.Sleep(1000);
                                        SendKeys.SendWait(Ruta);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(1000);
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Adjuntar Archivo", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Click("//a[contains(@id,'btnActualizar')]");
                                        Thread.Sleep(2000);
                                        selenium.AcceptAlert();
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Registro Actualizado", true, file);
                                        //ELIMINAR
                                        Thread.Sleep(6000);
                                        selenium.Click("//*[@id='tblBiEdfor']/tbody/tr/td[5]/a");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Registro a Eliminar",true, file);
                                        Thread.Sleep(2000);
                                        selenium.Click("//a[contains(@id,'btnEliminar')]");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Registro Eliminado", true, file);
                                    }

                                    //////
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
        public void BI_EducaciónNoFormal()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_EducaciónNoFormal")
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
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["EnFModalidad"].ToString().Length != 0 && rows["EnFModalidad"].ToString() != null &&
                                rows["EnFNomEstudions"].ToString().Length != 0 && rows["EnFNomEstudions"].ToString() != null &&
                                rows["EnFAnoInicio"].ToString().Length != 0 && rows["EnFAnoInicio"].ToString() != null &&
                                rows["EnFInstitucion"].ToString().Length != 0 && rows["EnFInstitucion"].ToString() != null &&
                                rows["EnFMesInicio"].ToString().Length != 0 && rows["EnFMesInicio"].ToString() != null &&
                                rows["EnFDiaInicio"].ToString().Length != 0 && rows["EnFDiaInicio"].ToString() != null &&
                                rows["EnFAnoFin"].ToString().Length != 0 && rows["EnFAnoFin"].ToString() != null &&
                                rows["EnFMesFin"].ToString().Length != 0 && rows["EnFMesFin"].ToString() != null &&
                                rows["EnFDiaFin"].ToString().Length != 0 && rows["EnFDiaFin"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["EnFTerminado"].ToString().Length != 0 && rows["EnFTerminado"].ToString() != null &&
                                rows["EnfTiempoEstudio"].ToString().Length != 0 && rows["EnfTiempoEstudio"].ToString() != null &&
                                rows["EnfUnidad"].ToString().Length != 0 && rows["EnfUnidad"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string EnFModalidad = rows["EnFModalidad"].ToString();
                                string EnFNomEstudions = rows["EnFNomEstudions"].ToString();
                                string EnFAnoInicio = rows["EnFAnoInicio"].ToString();
                                string EnFInstitucion = rows["EnFInstitucion"].ToString();
                                string EnFMesInicio = rows["EnFMesInicio"].ToString();
                                string EnFDiaInicio = rows["EnFDiaInicio"].ToString();
                                string EnFAnoFin = rows["EnFAnoFin"].ToString();
                                string EnFMesFin = rows["EnFMesFin"].ToString();
                                string EnFDiaFin = rows["EnFDiaFin"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string EnFTerminado = rows["EnFTerminado"].ToString();
                                string EnfTiempoEstudio = rows["EnfTiempoEstudio"].ToString();
                                string EnfUnidad = rows["EnfUnidad"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();
                                string Ruta = rows["Ruta"].ToString();

                                try
                                {

                                    string database = "";
                                    string user = "";
                                    string CodEmpr = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                        CodEmpr = "9";

                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                        CodEmpr = "421";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                        CodEmpr = "421";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                        CodEmpr = "9";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //limpiar procesos
                                    Process[] processes = Process.GetProcessesByName("chromedriver");
                                    if (processes.Length > 0)
                                    {
                                        for (int i = 0; i < processes.Length; i++)
                                        {
                                            processes[i].Kill();
                                        }
                                    }
                                    //REGISTROS
                                    string eliminarRegistro1 = $"delete BI_DTPNF where COD_EMPR='{CodEmpr}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRegistro1, database, user);
                                    string eliminarRegistro = $"delete bi_ednfo where cod_empl='{EmpleadoUser}' and act_usua='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                    //INICIO
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(1000);
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Mi Educación No Formal')]");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Educación no formal", true, file);
                                    //MODALIDAD
                                    selenium.SelectElementByName("//select[contains(@name,'ddlNomModi')]", EnFModalidad);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Modalidad", true, file);
                                    //ESTUDIOS
                                    selenium.SelectElementByName("//select[contains(@name,'ddlNomEstu')]", EnFNomEstudions);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Estudios", true, file);
                                    //INSTITUCION
                                    selenium.SendKeys("//input[contains(@name,'txtNomInst')]", EnFInstitucion);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Institución", true, file);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(3000);
                                    selenium.Scroll("//a[@id='ctl00_ContenidoPagina_kcfFecInic_imgCalendario']/span");
                                    Thread.Sleep(3000);
                                    //FECHA INICIO
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_kcfFecInic_imgCalendario']/span");
                                    Thread.Sleep(3000);
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_kcfFecInic_cetxtFecha_day_0_0']");
                                    Thread.Sleep(3000);
                                    //FECHA FIN
                                    Thread.Sleep(3000);
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_kcfFecTerm_imgCalendario']/span");
                                    Thread.Sleep(3000);
                                    selenium.Click("//div[@id='ctl00_ContenidoPagina_kcfFecTerm_cetxtFecha_day_0_0']");
                                    Thread.Sleep(3000);
                                    selenium.Screenshot("Fechas", true, file);
                                    Thread.Sleep(500);
                                    //Validación 1 País
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", ValPais);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    Thread.Sleep(500);
                                   
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    Thread.Sleep(500);
                                    
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                   
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //TERMINO ESTUDIOS
                                    
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[contains(@name,'ddlEstTerm')]", EnFTerminado);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Termino estudios", true, file);
                                    //TIEMPO ESTUDIO
                                    selenium.Scroll("//input[contains(@name,'txtTieEstu')]");
                                    selenium.SendKeys("//input[contains(@name,'txtTieEstu')]", EnfTiempoEstudio);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Tiempo estudio", true, file);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");

                                    if (database == "SQL")
                                    {
                                        selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_ddlUniTiem']", EnfUnidad);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Unidad estudios", true, file);
                                    }
                                    //GUARDAR
                                    selenium.Click("//a[@id='btnGuardar']/span");
                                    Thread.Sleep(4000);
                                    selenium.Screenshot("Registrada educación no formal", true, file);
                                    //ACTUALIZAR 
                                    if (selenium.ExistControl("//*[@id='tblBiEdnfo']/tbody/tr/td[5]/a"))
                                    {
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Existe educación no formal", true, file);
                                        selenium.Click("//*[@id='tblBiEdnfo']/tbody/tr/td[5]/a");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Actualizar educación no formal", true, file);
                                        //CIUDAD
                                        selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                        Thread.Sleep(2000);
                                        selenium.Click("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]");
                                        Thread.Sleep(500);
                                        selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipUbi_txtDivPoli')]", Ciudad);
                                        Thread.Sleep(500);
                                        SendKeys.SendWait("{TAB}");
                                        selenium.Click("//div[@id='ctl00_pBotones']/div");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Ciudad", true, file);
                                        ////ADJUNTO
                                        selenium.Scroll("//*[@id=\"ctl00_ContenidoPagina_lblAdjunto\"]");
                                        Thread.Sleep(1000);
                                        selenium.Click("//*[@id=\"ctl00_ContenidoPagina_lblAdjunto\"]");
                                        Thread.Sleep(1000);
                                        SendKeys.SendWait(Ruta);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(1000);
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Adjuntar Archivo", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Click("//a[contains(@id,'btnActualizar')]");
                                        Thread.Sleep(2000);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Registro Actualizado", true, file);
                                        //ELIMINAR
                                        Thread.Sleep(2000);
                                        selenium.Click("//*[@id='tblBiEdnfo']/tbody/tr/td[5]/a");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Registro a Eliminar", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Click("//a[contains(@id,'btnEliminar')]");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Registro Eliminado", true, file);

                                    }
                                    else
                                    {
                                        Thread.Sleep(500);
                                        selenium.Screenshot("No existe educacion no formal", true, file);

                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se inserto el estudio no formal");
                                    }
                                    Thread.Sleep(2000);
                                    selenium.Close();
                                    //////
                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
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
        public void BI_ExperienciaLaboral()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_ExperienciaLaboral")
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
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["ElEmpresa"].ToString().Length != 0 && rows["ElEmpresa"].ToString() != null &&
                                rows["ElDireccion"].ToString().Length != 0 && rows["ElDireccion"].ToString() != null &&
                                rows["ElTelefono"].ToString().Length != 0 && rows["ElTelefono"].ToString() != null &&
                                rows["ElTipEmpresa"].ToString().Length != 0 && rows["ElTipEmpresa"].ToString() != null &&
                                rows["ElEmplActual"].ToString().Length != 0 && rows["ElEmplActual"].ToString() != null &&
                                rows["ElCargoEjecutivo"].ToString().Length != 0 && rows["ElCargoEjecutivo"].ToString() != null &&
                                rows["ElDedicacion"].ToString().Length != 0 && rows["ElDedicacion"].ToString() != null &&
                                rows["ElFecIngAno"].ToString().Length != 0 && rows["ElFecIngAno"].ToString() != null &&
                                rows["ElFecIngMes"].ToString().Length != 0 && rows["ElFecIngMes"].ToString() != null &&
                                rows["ElFecIniDia"].ToString().Length != 0 && rows["ElFecIniDia"].ToString() != null &&
                                rows["ElFecRetAno"].ToString().Length != 0 && rows["ElFecRetAno"].ToString() != null &&
                                rows["ElFecRetMes"].ToString().Length != 0 && rows["ElFecRetMes"].ToString() != null &&
                                rows["ElFecRetDia"].ToString().Length != 0 && rows["ElFecRetDia"].ToString() != null &&
                                rows["ElManejaPersonal"].ToString().Length != 0 && rows["ElManejaPersonal"].ToString() != null &&
                                rows["ElCargoDesempeno"].ToString().Length != 0 && rows["ElCargoDesempeno"].ToString() != null &&
                                rows["ElArea"].ToString().Length != 0 && rows["ElArea"].ToString() != null &&
                                rows["ELTipContrato"].ToString().Length != 0 && rows["ELTipContrato"].ToString() != null &&
                                rows["ELMotivoRetiro"].ToString().Length != 0 && rows["ELMotivoRetiro"].ToString() != null &&
                                rows["Ciudad"].ToString().Length != 0 && rows["Ciudad"].ToString() != null &&
                                rows["ValPais"].ToString().Length != 0 && rows["ValPais"].ToString() != null &&
                                rows["ValDepartamento"].ToString().Length != 0 && rows["ValDepartamento"].ToString() != null &&
                                rows["ValCaracteresEsp"].ToString().Length != 0 && rows["ValCaracteresEsp"].ToString() != null &&
                                rows["ElJefeInmediato"].ToString().Length != 0 && rows["ElJefeInmediato"].ToString() != null &&
                                rows["ELCarogoJefeInme"].ToString().Length != 0 && rows["ELCarogoJefeInme"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string ElEmpresa = rows["ElEmpresa"].ToString();
                                string ElDireccion = rows["ElDireccion"].ToString();
                                string ElTelefono = rows["ElTelefono"].ToString();
                                string ElTipEmpresa = rows["ElTipEmpresa"].ToString();
                                string ElEmplActual = rows["ElEmplActual"].ToString();
                                string ElCargoEjecutivo = rows["ElCargoEjecutivo"].ToString();
                                string ElDedicacion = rows["ElDedicacion"].ToString();
                                string ElFecIngAno = rows["ElFecIngAno"].ToString();
                                string ElFecIngMes = rows["ElFecIngMes"].ToString();
                                string ElFecIniDia = rows["ElFecIniDia"].ToString();
                                string ElFecRetAno = rows["ElFecRetAno"].ToString();
                                string ElFecRetMes = rows["ElFecRetMes"].ToString();
                                string ElFecRetDia = rows["ElFecRetDia"].ToString();
                                string ElManejaPersonal = rows["ElManejaPersonal"].ToString();
                                string ElCargoDesempeno = rows["ElCargoDesempeno"].ToString();
                                string ElArea = rows["ElArea"].ToString();
                                string ELTipContrato = rows["ELTipContrato"].ToString();
                                string ELMotivoRetiro = rows["ELMotivoRetiro"].ToString();
                                string Ciudad = rows["Ciudad"].ToString();
                                string ValPais = rows["ValPais"].ToString();
                                string ValDepartamento = rows["ValDepartamento"].ToString();
                                string ValCaracteresEsp = rows["ValCaracteresEsp"].ToString();
                                string ElJefeInmediato = rows["ElJefeInmediato"].ToString();
                                string ELCarogoJefeInme = rows["ELCarogoJefeInme"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string Ruta = rows["Ruta"].ToString();
                                string TipoDocumento = rows["TipoDocumento"].ToString();

                                try
                                {

                                    string database = "";
                                    string user = "";
                                    string CodEmpr = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                        CodEmpr = "9";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                        CodEmpr = "421";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                        CodEmpr = "421";

                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                        CodEmpr = "9";

                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //limpiar procesos
                                    Process[] processes = Process.GetProcessesByName("chromedriver");
                                    if (processes.Length > 0)
                                    {
                                        for (int i = 0; i < processes.Length; i++)
                                        {
                                            processes[i].Kill();
                                        }
                                    }
                                    //BORRA REGISTROS
                                    string eliminarRegistro1 = $"DELETE BI_ARHEX where cod_empr='{CodEmpr}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRegistro1, database, user);
                                    Thread.Sleep(2000);
                                    string eliminarRegistro2 = $"DELETE BI_DTPHV where cod_empr='{CodEmpr}' and ACT_USUA='{EmpleadoUser}'";
                                    db.UpdateDeleteInsert(eliminarRegistro2, database, user);
                                    string eliminarRegistro3 = $"DELETE bi_hvext where COD_EMPL='{EmpleadoUser}' and cod_empr='{CodEmpr}'";
                                    db.UpdateDeleteInsert(eliminarRegistro3, database, user);
                                    //INICIO
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(1000);
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Mi Experiencia Laboral')]");
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Experiencia Laboral", true, file);
                                    //EMPRESA
                                    selenium.SendKeys("//input[contains(@id,'txtNomEmpr')]", ElEmpresa);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Empresa ", true, file);
                                    //DIRECCION
                                    selenium.SendKeys("//input[contains(@id,'txtDirEmpr')]", ElDireccion);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Dirección", true, file);
                                    //TELEFONO
                                    selenium.SendKeys("//input[contains(@id,'txtTelEmpr')]", ElTelefono);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Teléfono ", true, file);
                                    //TIPO EMPRESA
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipEmpr')]", ElTipEmpresa);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Tipo empresa ", true, file);
                                    //EMPLEO ACTUAL
                                    selenium.SelectElementByName("//select[contains(@id,'ddlEmpActu')]", ElEmplActual);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Empleo actual ", true, file);
                                    //CARGO
                                    selenium.SelectElementByName("//select[contains(@id,'ddlCarEjec')]", ElCargoEjecutivo);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Cargo Laboral", true, file);
                                    //DEDICACION
                                    selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlDedIcac']", ElDedicacion);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Dedicación", true, file);
                                    //FECHA INICIO
                                    string FechaInicio = $"{ElFecIniDia}/{ElFecIngMes}/{ElFecIngAno}";
                                 
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_kcfFecIngr_txtFecha']", FechaInicio);
                                    selenium.Screenshot("Fecha inicio", true, file);
                                    //FECHA RETIRO
                                    string FechaRetiro = $"{ElFecRetDia}/{ElFecRetMes}/{ElFecRetAno}";
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(2000);
                                    selenium.SendKeys("//input[@id='ctl00_ContenidoPagina_kcfFecReti_txtFecha']", FechaRetiro);
                                    //MANEJA PERSONAL
                                    selenium.SelectElementByName("//select[contains(@name,'ddlManPers')]", ElManejaPersonal);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Maneja Personal", true, file);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.ScrollTo("0", "250");
                                    Thread.Sleep(1500);
                                    //CARGO DESEMPEÑADO
                                    selenium.SendKeys("//input[contains(@id,'txtCarDese')]", ElCargoDesempeno);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Cargo", true, file);
                                    //AREA
                                    selenium.SendKeys("//input[contains(@id,'txtDepEmpr')]", ElArea);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Área", true, file);
                                    //TIPO CONTRATO
                                    selenium.SelectElementByName("//select[contains(@id,'ddlTipCont')]", ELTipContrato);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Tipo contrato ", true, file);
                                    //MOTIVO RETIRO
                                    selenium.SelectElementByName("//select[contains(@id,'ddlMotReti')]", ELMotivoRetiro);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Retiro", true, file);
                                    selenium.ScrollTo("0", "250");
                                    //Validación 1 País
                                    selenium.Scroll("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]", ValPais);
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Ingreso de País para validar", true, file);
                                    Thread.Sleep(1500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación 2 Departamento
                                    Thread.Sleep(500);

                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]", ValDepartamento);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Caracteres especiales
                                    Thread.Sleep(500);
                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]", ValCaracteresEsp);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Ingreso de Departamento para validar", true, file);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    //Validación Exitosa
                                    Thread.Sleep(500);

                                    selenium.SendKeys("//input[contains(@id,'ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli')]", Ciudad);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{TAB}");
                                    Thread.Sleep(1000);
                                    selenium.Screenshot("Validación de campo en blanco", true, file);
                                    Thread.Sleep(1000);
                                    //JEFE INMEDIATO
                                    selenium.Scroll("//input[contains(@id,'txtJefInme')]");
                                    selenium.SendKeys("//input[contains(@id,'txtJefInme')]", ElJefeInmediato);
                                    Thread.Sleep(1000);
                                    //CARGO JEFE INMEDIATO
                                    selenium.SendKeys("//input[contains(@id,'txtCarJefe')]", ELCarogoJefeInme);
                                    Thread.Sleep(500);
                                    selenium.Screenshot("Datos Experiencia Laboral", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(2000);
                                    selenium.ChangeAuxWindow();
                                    selenium.Click("//input[@id='ctl00_ContentPopapModel_btnNoProLog']");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Registro Exitoso", true, file);
                                    Thread.Sleep(2000);
                                    selenium.AcceptAlert();
                                    Thread.Sleep(2000);
                                    //VERIFICAR REGISTRO
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(200);
                                    selenium.Click("//a[contains(.,'Mi Experiencia Laboral')]");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Guardar Experiencia Laboral", true, file);
                                    Thread.Sleep(2000);
                                    //ACTUALIZAR REGISTRO
                                    if (selenium.ExistControl("//*[@id='ctl00_ContenidoPagina_dgrBiHvext_ctl03_LinkButton1']"))
                                    {
                                        Thread.Sleep(1500);
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_dgrBiHvext_ctl03_LinkButton1']");
                                        Thread.Sleep(2000);
                                        selenium.Scroll("//*[@id='ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli']");
                                        Thread.Sleep(1000);
                                        //CIUDAD
                                        Thread.Sleep(2000);
                                        selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KCtrDivipEmpre_txtDivPoli']", Ciudad);
                                        Thread.Sleep(500);
                                        //ADJUNTO
                                        selenium.Scroll("//*[@id=\"ctl00_ContenidoPagina_lblAdjunto\"]");
                                        Thread.Sleep(1000);
                                        selenium.Click("//*[@id=\"ctl00_ContenidoPagina_lblAdjunto\"]");
                                        Thread.Sleep(2000);
                                        SendKeys.SendWait(Ruta);
                                        Thread.Sleep(2000);
                                        SendKeys.SendWait("{ENTER}");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Adjunto", true, file);
                                        Thread.Sleep(2000);
                                        selenium.SelectElementByName("//select[contains(@id,'ctl00_ContenidoPagina_KCtrTipoDocumento_ddlTIP_DOCU')]", TipoDocumento);
                                        Thread.Sleep(1000);
                                        selenium.Screenshot("Adjuntar Archivo", true, file);
                                        Thread.Sleep(2000);
                                        selenium.Click("//a[contains(@id,'btnActualizar')]");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Experiencia laboral editada", true, file);
                                        //ELIMINAR REGISTRO
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_dgrBiHvext_ctl03_LinkButton1']");
                                        Thread.Sleep(2000);
                                        selenium.Click("//a[contains(@id,'btnEliminar')]");
                                        Thread.Sleep(500);
                                        selenium.Screenshot("Elimina Experiencia Laboral", true, file);
                                    }
                                    else
                                    {
                                        Thread.Sleep(500);
                                        selenium.Screenshot("No Existe Exp Laboral", true, file);

                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se inserto la experiencia laboral");
                                    }
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
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
        public void BI_Familiares()
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

                if (methodname.Replace(" ", string.Empty) == "Web_Kactus_Test_V2.Modulo_BI.BI_Familiares")
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
                                rows["EmpleadoUser"].ToString().Length != 0 && rows["EmpleadoUser"].ToString() != null &&
                                rows["EmpleadoPass"].ToString().Length != 0 && rows["EmpleadoPass"].ToString() != null &&
                                rows["FaIdenti"].ToString().Length != 0 && rows["FaIdenti"].ToString() != null &&
                                rows["FaNombre"].ToString().Length != 0 && rows["FaNombre"].ToString() != null &&
                                rows["FaApellido"].ToString().Length != 0 && rows["FaApellido"].ToString() != null &&
                                rows["FaSexo"].ToString().Length != 0 && rows["FaSexo"].ToString() != null &&
                                rows["FaNaciAno"].ToString().Length != 0 && rows["FaNaciAno"].ToString() != null &&
                                rows["FaNaciMes"].ToString().Length != 0 && rows["FaNaciMes"].ToString() != null &&
                                rows["FaNaciDia"].ToString().Length != 0 && rows["FaNaciDia"].ToString() != null &&
                                rows["FaVive"].ToString().Length != 0 && rows["FaVive"].ToString() != null &&
                                rows["FaGrupoSangre"].ToString().Length != 0 && rows["FaGrupoSangre"].ToString() != null &&
                                rows["FaFactosSangui"].ToString().Length != 0 && rows["FaFactosSangui"].ToString() != null &&
                                rows["FaEstadoCivil"].ToString().Length != 0 && rows["FaEstadoCivil"].ToString() != null &&
                                rows["url"].ToString().Length != 0 && rows["url"].ToString() != null &&
                                rows["Maquina"].ToString().Length != 0 && rows["Maquina"].ToString() != null
                                )
                            {
                                string EmpleadoUser = rows["EmpleadoUser"].ToString();
                                string EmpleadoPass = rows["EmpleadoPass"].ToString();
                                string FaIdenti = rows["FaIdenti"].ToString();
                                string FaNombre = rows["FaNombre"].ToString();
                                string FaApellido = rows["FaApellido"].ToString();
                                string FaSexo = rows["FaSexo"].ToString();
                                string FaNaciAno = rows["FaNaciAno"].ToString();
                                string FaNaciMes = rows["FaNaciMes"].ToString();
                                string FaNaciDia = rows["FaNaciDia"].ToString();
                                string FaVive = rows["FaVive"].ToString();
                                string FaGrupoSangre = rows["FaGrupoSangre"].ToString();
                                string FaFactosSangui = rows["FaFactosSangui"].ToString();
                                string FaEstadoCivil = rows["FaEstadoCivil"].ToString();
                                string url = rows["url"].ToString();
                                string Maquina = rows["Maquina"].ToString();
                                string FaEstadoCivil1 = rows["FaEstadoCivil1"].ToString();
                                string Avenida = rows["Avenida"].ToString();
                                string numero = rows["numero"].ToString();
                                string letra = rows["letra"].ToString();
                                string bis = rows["bis"].ToString();
                                string letra2 = rows["letra2"].ToString();
                                string ubicacion = rows["ubicacion"].ToString();
                                string numero1 = rows["numero1"].ToString();
                                string letra3 = rows["letra3"].ToString();
                                string numero2 = rows["numero2"].ToString();
                                string ubicacion1 = rows["ubicacion1"].ToString();
                                try
                                {

                                    string database = "";
                                    string user = "";
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestnewverauto/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    if (url.ToLower() == "http://dwtfskscm/selfservicetestoranewverauto/".ToLower())
                                    {
                                        database = "ORA";
                                        user = "ODESAR";
                                    }

                                    if (url.ToLower() == "http://dwtfsk:8094/".ToLower())
                                    {

                                        database = "ORA";
                                        user = "ODESAR";
                                    }
                                    if (url.ToLower() == "http://dwtfsk:8093/".ToLower())
                                    {
                                        database = "SQL";
                                        user = "SDesar";
                                    }
                                    string[] split = methodname.Split('.');
                                    string nombre = split[2];
                                    string[] split1 = nombre.Split('_');
                                    string modulo = split1[0];
                                    string file = fv.CrearDocumentoWordDinamico(app, database, modulo, nombre + CaseId);
                                    //limpiar procesos
                                    Process[] processes = Process.GetProcessesByName("chromedriver");
                                    if (processes.Length > 0)
                                    {
                                        for (int i = 0; i < processes.Length; i++)
                                        {
                                            processes[i].Kill();
                                        }
                                    }
                                    string eliminarRegistro = $"DELETE from bi_famil where cod_empl='{EmpleadoUser}' and cod_fami='{FaIdenti}'";
                                    db.UpdateDeleteInsert(eliminarRegistro, database, user);
                                    string eliminarRegistro1 = $"DELETE from bi_famil where cod_empl='{EmpleadoUser}' and cod_fami='880804045440'";
                                    db.UpdateDeleteInsert(eliminarRegistro1, database, user); 
                                    //INICIO PRUEBA
                                    selenium.LoginApps(app, EmpleadoUser, EmpleadoPass, url, file);
                                    Thread.Sleep(1000);
                                    //MI INFORMACION PERSONAL/FAMILIARES
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Mis Familiares')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Familiares", true, file);
                                    //NUEVO
                                    bool existe = selenium.ExistControl("//a[@id='ctl00_btnNuevo']");
                                    if (existe)
                                    {
                                        selenium.Click("//a[@id='ctl00_btnNuevo']");
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Registro Nuevo", true, file);
                                    }
                                    //IDENTIFICACION
                                    selenium.Click("//input[contains(@id,'txtCodFami')]");
                                    selenium.SendKeys("//input[contains(@id,'txtCodFami')]", FaIdenti);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Identificación", true, file);
                                    //NOMBRE
                                    selenium.Click("//input[contains(@id,'txtNomFami1')]");
                                    selenium.SendKeys("//input[contains(@id,'txtNomFami1')]", FaNombre);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Nombre", true, file);
                                    //APELLIDO
                                    selenium.Click("//input[contains(@id,'txtApeFami1')]");
                                    selenium.SendKeys("//input[contains(@id,'txtApeFami1')]", FaApellido);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Apellido", true, file);
                                    selenium.Click("//div[@id='ctl00_pBotones']/div");
                                    selenium.ScrollTo("0", "250");
                                    Thread.Sleep(1000);
                                    //SEXO
                                    selenium.SelectElementByName("//select[contains(@id,'ddlSexFami')]", FaSexo);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Sexo", true, file);
                                    //FECHA NACIMIENTO
                                    string FechaNacimiento = $"{FaNaciDia}/{FaNaciMes}/{FaNaciAno}";
                                    selenium.Click("//input[contains(@id,'kcfFecNaci')]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//input[contains(@id,'kcfFecNaci')]", FechaNacimiento);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Fecha Nacimiento", true, file);
                                    //GRUPO SANGUINEO
                                    bool valGrupSang = selenium.ExistControl("//select[contains(@id,'ddlGruSang')]");
                                    if (valGrupSang)
                                    {
                                        selenium.Click("//select[contains(@id,'ddlGruSang')]");
                                        Thread.Sleep(2000);
                                        selenium.SelectElementByName("//select[contains(@id,'ddlGruSang')]", FaGrupoSangre);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Grupo Sangre", true, file);
                                        selenium.Click("//select[contains(@id,'ddlFacSang')]");
                                        Thread.Sleep(2000);
                                        selenium.SelectElementByName("//select[contains(@id,'ddlFacSang')]", FaFactosSangui);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Sangre", true, file);

                                    }

                                    //ESTADO CIVIL
                                    bool valEstCivil = selenium.ExistControl("//select[contains(@id,'ddlEstCivi')]");
                                    if (valEstCivil)
                                    {
                                        selenium.Click("//select[contains(@id,'ddlEstCivi')]");
                                        Thread.Sleep(1500);
                                        selenium.SelectElementByName("//select[contains(@id,'ddlEstCivi')]", FaEstadoCivil);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Estado Civil", true, file);
                                    }
                                    else
                                    {
                                        selenium.Click("//*[@id='ctl00_ContenidoPagina_ddlEstCivi']");
                                        Thread.Sleep(1500);
                                        selenium.SelectElementByName("//*[@id='ctl00_ContenidoPagina_ddlEstCivi']", FaEstadoCivil);
                                        Thread.Sleep(2000);
                                        selenium.Screenshot("Estado Civil", true, file);
                                    }
                                    Thread.Sleep(1000);
                                    selenium.Click("//*[@id='ctl00_pBotones']/div[1]");
                                    //CIUDAD
                                    selenium.Scroll("//*[@id='ctl00_ContenidoPagina_KCtrDivip_txtDivPoli']");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KCtrDivip_txtDivPoli']", "COLOMBIA, CUNDINAMARCA, BOGOTA D.C.");
                                    Thread.Sleep(1500);
                                    selenium.Click("//*[@id='ctl00_pBotones']/div[1]");
                                    Thread.Sleep(1500);
                                    selenium.SendKeys("//*[@id='ctl00_ContenidoPagina_KCtrDivip_txtDivPoli']", "COLOMBIA, CUNDINAMARCA, BOGOTA D.C.");
                                    selenium.Screenshot("Ciudad", true, file);
                                    //DIRECCION
                                    selenium.Click("//a[@id='ctl00_ContenidoPagina_lkbDirEcci']/span");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Direccion", true, file);
                                    Thread.Sleep(2000);
                                    selenium.SelectElementByName("//select[@id='ctl00_ContentPopapModel_KCtrlDireccion1_Vias']", Avenida);
                                    selenium.SendKeys("//input[@id='inpnumero1']", numero);
                                    selenium.SelectElementByName("//select[@id='ddlletra1']", letra);
                                    selenium.SelectElementByName("//select[@id='ddlbis']", bis);
                                    selenium.SelectElementByName("//select[@id='ddlletra2']", letra2);
                                    selenium.SelectElementByName("//select[@id='ddlubi1']", ubicacion);
                                    selenium.SendKeys("//input[@id='inpnumero2']", numero1);
                                    selenium.SelectElementByName("//select[@id='ddlletra3']", letra3);
                                    selenium.SendKeys("//input[@id='inpnumero3']", numero2);
                                    selenium.SelectElementByName("//select[@id='ddlubi2']", ubicacion1);
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Direccion Registrada", true, file);
                                    Thread.Sleep(2000);
                                    selenium.Click("//input[@id='ctl00_ContentPopapModel_btnterminar']");
                                    Thread.Sleep(2000);
                                    selenium.Screenshot("Direccion Aplicada", true, file);
                                    //GUARDAR
                                    selenium.Click("//a[contains(@id,'btnGuardar')]");
                                    Thread.Sleep(5000);
                                    selenium.Screenshot("Guardar Familiares", true, file);
                                    Thread.Sleep(1000);
                                    if(selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                    {
                                        selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                    }
                                    //MI INFORMACION PERSONAL/FAMILIARES
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Mis Familiares')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Familiares", true, file);
                                    //EDITAR REGISTRO
                                    if (selenium.ExistControl("//a[@id='ctl00_ContenidoPagina_Familiares_ctl03_LinkButton1']/i"))
                                    {
                                        if (database == "SQL")
                                        {
                                            Thread.Sleep(1500);
                                            selenium.Screenshot("Familiar a editar", true, file);
                                            selenium.Click("//a[@id='ctl00_ContenidoPagina_Familiares_ctl03_LinkButton1']/i");
                                            Thread.Sleep(1500);
                                            selenium.Scroll("//select[@id='ctl00_ContenidoPagina_ddlEstCivi']");
                                            Thread.Sleep(1500);
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlEstCivi']", FaEstadoCivil1);
                                            Thread.Sleep(1500);
                                            selenium.Screenshot("Estado Civil editado", true, file);
                                            Thread.Sleep(3000);
                                            selenium.Click("//a[@id='btnActualizar']");
                                            Thread.Sleep(6000);
                                            if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                            {
                                                selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                            }
                                        }
                                        else
                                        {
                                            Thread.Sleep(1500);
                                            selenium.Screenshot("Familiar a editar", true, file);
                                            selenium.Click("//a[@id='ctl00_ContenidoPagina_Familiares_ctl04_LinkButton1']/i");
                                            Thread.Sleep(1500);
                                            selenium.SelectElementByName("//select[@id='ctl00_ContenidoPagina_ddlEstCivi']", FaEstadoCivil1);
                                            Thread.Sleep(1500);
                                            selenium.Screenshot("Estado Civil editado", true, file);
                                            Thread.Sleep(3000);
                                            selenium.Click("//a[@id='btnActualizar']");
                                            Thread.Sleep(6000);
                                            if (selenium.ExistControl("/html/body/div[1]/div/div[4]/div/button"))
                                            {
                                                selenium.Click("/html/body/div[1]/div/div[4]/div/button");
                                            }
                                        }
                                        

                                        try
                                        {
                                            Screenshot("Registro Actualizado", true, file);
                                            selenium.AcceptAlert();
                                        }
                                        catch
                                        {
                                        }
                                    }
                                    else
                                    {
                                        selenium.Screenshot("No existe Familiar para editar", true, file);

                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se inserto la referencia familiar");
                                    }

                                    //MI INFORMACION PERSONAL/FAMILIARES
                                    selenium.Click("//a[contains(.,'MI INFORMACIÓN PERSONA')]");
                                    Thread.Sleep(2000);
                                    selenium.Click("//a[contains(.,'Mis Familiares')]");
                                    Thread.Sleep(1500);
                                    selenium.Screenshot("Familiares", true, file);
                                    //ELIMINAR REGISTRO
                                    Thread.Sleep(4000);
                                    if (selenium.ExistControl("//*[@id='ctl00_ContenidoPagina_Familiares_ctl03_LinkButton1']/i"))
                                    {
                                        Thread.Sleep(1500);
                                        selenium.Screenshot("Familiar", true, file);
                                        if (database == "SQL")
                                        {
                                            selenium.Click("//*[@id='ctl00_ContenidoPagina_Familiares_ctl03_LinkButton1']/i");
                                            selenium.Click("//a[contains(@id,'btnEliminar')]");
                                            Thread.Sleep(1500);
                                            selenium.Screenshot("Elimina Familiar", true, file);
                                        }
                                        else
                                        {
                                            selenium.Click("//*[@id='ctl00_ContenidoPagina_Familiares_ctl04_LinkButton1']/i");
                                            selenium.Click("//a[contains(@id,'btnEliminar')]");
                                            Thread.Sleep(1500);
                                            selenium.Screenshot("Elimina Familiar", true, file);
                                        }
                                        

                                        try
                                        {
                                            Screenshot("Registro Eliminado", true, file);
                                            selenium.AcceptAlert();
                                        }
                                        catch
                                        {
                                        }
                                    }
                                    else
                                    {
                                        selenium.Screenshot("No existe Familiar", true, file);

                                        errorMessagesMetodo.Add(" ::::::::::::::::::::::" + "MSG: No se inserto la referencia familiar");
                                    }
                                    //////
                                    selenium.Close();
                                    fv.ConvertWordToPDF(file, database);
                                    ////////////////////////////////////////////////////
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

