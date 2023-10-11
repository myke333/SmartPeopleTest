namespace Web_Kactus_Test.UIMap_RL_New_VersionClasses
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

    public partial class UIMap_RL_New_Version:FuncionesVitales
    {

        /// <summary>
        /// AddArchivoRlFamil2: use 'AddArchivoRlFamil2Params' para pasar parámetros a este método.
        /// </summary>
        public void AddArchivoRlFamil2(string file, string Documentos)
        {
            #region Variable Declarations
            WinControl uIChromeLegacyWindowDocument = this.UIKactusRLReclutamientWindow.UIChromeLegacyWindowWindow.UIChromeLegacyWindowDocument;
            WinComboBox uINombreComboBox = this.UIAbrirWindow.UIItemWindow.UINombreComboBox;
            WinButton uIAbrirButton = this.UIAbrirWindow.UIAbrirWindow1.UIAbrirButton;
            #endregion

            //// Clic 'Chrome Legacy Window' documento
            //Mouse.Click(uIChromeLegacyWindowDocument, new Point(393, 549));

            //// Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
            //uINombreComboBox.EditableItem = this.AddArchivoRlFamil2Params.UINombreComboBoxEditableItem;

            //// Clic '&Abrir' botón
            //Mouse.Click(uIAbrirButton, new Point(34, 14));
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

        public virtual AddArchivoRlFamil2Params AddArchivoRlFamil2Params
        {
            get
            {
                if ((this.mAddArchivoRlFamil2Params == null))
                {
                    this.mAddArchivoRlFamil2Params = new AddArchivoRlFamil2Params();
                }
                return this.mAddArchivoRlFamil2Params;
            }
        }

        private AddArchivoRlFamil2Params mAddArchivoRlFamil2Params;
    }
    /// <summary>
    /// Parámetros que se van a pasar a 'AddArchivoRlFamil2'
    /// </summary>
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class AddArchivoRlFamil2Params
    {

        #region Fields
        /// <summary>
        /// Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
        /// </summary>
        public string UINombreComboBoxEditableItem = "ArchivoPrueba.pdf";
        #endregion
    }
}
