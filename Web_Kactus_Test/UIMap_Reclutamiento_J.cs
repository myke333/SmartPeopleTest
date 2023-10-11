namespace Web_Kactus_Test.UIMap_Reclutamiento_JClasses
{
    using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
    using System;
    using System.Collections.Generic;
    using System.CodeDom.Compiler;
    using Microsoft.VisualStudio.TestTools.UITest.Extension;
    using Microsoft.VisualStudio.TestTools.UITesting;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
    using Mouse = Microsoft.VisualStudio.TestTools.UITesting.Mouse;
    using MouseButtons = System.Windows.Forms.MouseButtons;
    using System.Drawing;
    using System.Windows.Input;
    using System.Text.RegularExpressions;
    using System.Threading;

    public partial class UIMap_Reclutamiento_J : FuncionesVitales
    {

        /// <summary>
        /// CargarCertificacionRlEmpdo: use 'CargarCertificacionRlEmpdoParams' para pasar parámetros a este método.
        /// </summary>
        public void CargarCertificacionRlEmpdo(string file, string Documentos)
        {
            #region Variable Declarations
            WinButton uIKactusRLReclutamientButton = this.UIKactusRLReclutamientWindow.UIItemPropertyPage.UIKactusRLReclutamientButton;
            WinComboBox uINombreComboBox = this.UIAbrirWindow.UIItemWindow.UINombreComboBox;
            WinButton uIAbrirButton = this.UIAbrirWindow.UIAbrirWindow1.UIAbrirButton;
            #endregion

            //// Clic '.:. Kactus RL Reclutamiento Web .:. - Google Chrom...' botón
            //Mouse.Click(uIKactusRLReclutamientButton, new Point(89, 22));

            //// Seleccionar 'gfhgffnbg' en cuadro combinado 'Nombre:'
            //uINombreComboBox.EditableItem = this.CargarCertificacionRlEmpdoParams.UINombreComboBoxEditableItem;

            //// Clic '&Abrir' botón
            //Mouse.Click(uIAbrirButton, new Point(42, 11));

            UITestControlCollection controlCol = this.UIAbrirWindow.UIItemWindow.GetChildren();
            for (int i = 0; i < controlCol.Count; i++)
            {
                if (controlCol[i].ControlType == "ComboBox")
                {
                    UITestControlCollection controlCol2 = controlCol[i].GetChildren();
                    for (int j = 0; j < controlCol2.Count; j++)
                    {
                        if (controlCol2[j].ClassName == "Edit")
                        {
                            UITestControlCollection controlCol3 = controlCol2[j].GetChildren();
                            for (int l = 0; l < controlCol3.Count; l++)
                            {
                                if (controlCol3[l].ControlType == "Edit")
                                {
                                    controlCol3[l].SetProperty("Text", Documentos);
                                }
                            }
                        }
                    }
                }
            }
            Thread.Sleep(500);
            Screenshot("Agregar archivo Self Service", true, file);

            try
            {
                Mouse.Click(this.UIAbrirWindow.UIAbrirWindow1.UIAbrirButton);
            }
            catch
            {
                Keyboard.SendKeys("{ENTER}");
            }
        }

        public virtual CargarCertificacionRlEmpdoParams CargarCertificacionRlEmpdoParams
        {
            get
            {
                if ((this.mCargarCertificacionRlEmpdoParams == null))
                {
                    this.mCargarCertificacionRlEmpdoParams = new CargarCertificacionRlEmpdoParams();
                }
                return this.mCargarCertificacionRlEmpdoParams;
            }
        }

        private CargarCertificacionRlEmpdoParams mCargarCertificacionRlEmpdoParams;

        /// <summary>
        /// AbrirDocumentos_Familiares
        /// </summary>
        public void AbrirDocumentos_Familiares()
        {
            #region Variable Declarations
            WinControl uIItemDropDownButton = this.UIAbrirWindow.UIAbrirSplitButton.UIItemDropDownButton;
            #endregion

            // Clic DropDownButton
            Mouse.Click(uIItemDropDownButton, new Point(123518971, 10));
        }

        /// <summary>
        /// CargarArchivos_Familiares
        /// </summary>
        public void CargarArchivos_Familiares()
        {
            #region Variable Declarations
            WinControl uIItemDropDownButton = this.UIAbrirWindow.UIAbrirSplitButton.UIItemDropDownButton;
            WinCustom uIKactusRLReclutamientCustom = this.UIKactusRLReclutamientWindow.UIItemPropertyPage.UIKactusRLReclutamientCustom;
            #endregion

            // Clic DropDownButton
            Mouse.Click(uIItemDropDownButton, new Point(163461439, 14));

            // Clic '.:. Kactus RL Reclutamiento Web .:. - Google Chrom...' control personalizado
            Mouse.Click(uIKactusRLReclutamientCustom, new Point(591, 106));
        }
    }
    /// <summary>
    /// Parámetros que se van a pasar a 'CargarCertificacionRlEmpdo'
    /// </summary>
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class CargarCertificacionRlEmpdoParams
    {

        #region Fields
        /// <summary>
        /// Seleccionar 'gfhgffnbg' en cuadro combinado 'Nombre:'
        /// </summary>
        public string UINombreComboBoxEditableItem = "gfhgffnbg";
        #endregion
    }
}
