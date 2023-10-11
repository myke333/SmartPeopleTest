namespace Web_Kactus_Test
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


    public partial class UIMap
    {

        /// <summary>
        /// ReportePDF: use 'ReportePDFParams' para pasar parámetros a este método.
        /// </summary>
        public void ReportePDF()
        {
            #region Variable Declarations
            WinComboBox uINombreComboBox = this.UIGuardarcomoWindow.UIPaneldedetallesPane.UINombreComboBox;
            WinButton uIGuardarButton = this.UIGuardarcomoWindow.UIGuardarWindow.UIGuardarButton;
            #endregion

            // Seleccionar 'ManualFunciones' en cuadro combinado 'Nombre:'
            uINombreComboBox.EditableItem = this.ReportePDFParams.UINombreComboBoxEditableItem;

            // Clic '&Guardar' botón
            Mouse.Click(uIGuardarButton, new Point(0, 6));
        }

        public virtual ReportePDFParams ReportePDFParams
        {
            get
            {
                if ((this.mReportePDFParams == null))
                {
                    this.mReportePDFParams = new ReportePDFParams();
                }
                return this.mReportePDFParams;
            }
        }

        private ReportePDFParams mReportePDFParams;

        /// <summary>
        /// Documentos1: use 'Documentos1Params' para pasar parámetros a este método.
        /// </summary>
        public void Documentos1()
        {
            #region Variable Declarations
            WinPane uIKACTUSSmartPeopleGooPane = this.UIKACTUSSmartPeopleGooWindow.UIKACTUSSmartPeopleGooPane;
            WinComboBox uINombreComboBox = this.UIAbrirWindow.UIItemWindow.UINombreComboBox;
            WinButton uIAbrirButton = this.UIAbrirWindow.UIAbrirWindow1.UIAbrirButton;
            #endregion

            // Clic 'KACTUS Smart People - Google Chrome' panel
            Mouse.Click(uIKACTUSSmartPeopleGooPane, new Point(1008, 643));

            // Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
            uINombreComboBox.EditableItem = this.Documentos1Params.UINombreComboBoxEditableItem;

            // Clic '&Abrir' botón
            Mouse.Click(uIAbrirButton, new Point(55, 14));
        }

        public virtual Documentos1Params Documentos1Params
        {
            get
            {
                if ((this.mDocumentos1Params == null))
                {
                    this.mDocumentos1Params = new Documentos1Params();
                }
                return this.mDocumentos1Params;
            }
        }

        private Documentos1Params mDocumentos1Params;

        /// <summary>
        /// Documentos2: use 'Documentos2Params' para pasar parámetros a este método.
        /// </summary>
        public void Documentos2()
        {
            #region Variable Declarations
            WinPane uIKACTUSSmartPeopleGooPane = this.UIKACTUSSmartPeopleGooWindow.UIKACTUSSmartPeopleGooPane;
            WinComboBox uINombreComboBox = this.UIAbrirWindow.UIItemWindow.UINombreComboBox;
            WinButton uIAbrirButton = this.UIAbrirWindow.UIAbrirWindow1.UIAbrirButton;
            #endregion

            // Clic 'KACTUS Smart People - Google Chrome' panel
            Mouse.Click(uIKACTUSSmartPeopleGooPane, new Point(1002, 688));

            // Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
            uINombreComboBox.EditableItem = this.Documentos2Params.UINombreComboBoxEditableItem;

            // Clic '&Abrir' botón
            Mouse.Click(uIAbrirButton, new Point(67, 14));
        }

        public virtual Documentos2Params Documentos2Params
        {
            get
            {
                if ((this.mDocumentos2Params == null))
                {
                    this.mDocumentos2Params = new Documentos2Params();
                }
                return this.mDocumentos2Params;
            }
        }

        private Documentos2Params mDocumentos2Params;
    }
    /// <summary>
    /// Parámetros que se van a pasar a 'ReportePDF'
    /// </summary>
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class ReportePDFParams
    {

        #region Fields
        /// <summary>
        /// Seleccionar 'ManualFunciones' en cuadro combinado 'Nombre:'
        /// </summary>
        public string UINombreComboBoxEditableItem = "ManualFunciones";
        #endregion
    }
    /// <summary>
    /// Parámetros que se van a pasar a 'Documentos1'
    /// </summary>
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class Documentos1Params
    {

        #region Fields
        /// <summary>
        /// Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
        /// </summary>
        public string UINombreComboBoxEditableItem = "ArchivoPrueba.pdf";
        #endregion
    }
    /// <summary>
    /// Parámetros que se van a pasar a 'Documentos2'
    /// </summary>
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class Documentos2Params
    {

        #region Fields
        /// <summary>
        /// Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
        /// </summary>
        public string UINombreComboBoxEditableItem = "ArchivoPrueba.pdf";
        #endregion
    }
}
