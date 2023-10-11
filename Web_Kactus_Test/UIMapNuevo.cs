namespace Web_Kactus_Test.UIMapNuevoClasses
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

    public partial class UIMapNuevo : FuncionesVitales
    {

        /// <summary>
        /// ArchivoPrueba2: use 'ArchivoPrueba2Params' para pasar parámetros a este método.
        /// </summary>
        public void ArchivoPrueba2(string file)
        {
            #region Variable Declarations
            WinPane uIKACTUSSmartPeopleGooPane = this.UIKACTUSSmartPeopleGooWindow.UIKACTUSSmartPeopleGooPane;
            WinList uIVistaElementosList = this.UIAbrirWindow.UIShellViewClient.UIVistaElementosList;
            WinButton uIAbrirButton = this.UIAbrirWindow.UIAbrirWindow1.UIAbrirButton;
            #endregion

            // Clic 'KACTUS Smart People - Google Chrome' panel
            Mouse.Click(uIKACTUSSmartPeopleGooPane, new Point(413, 680));

            // Seleccionar '' en cuadro de lista 'Vista Elementos'
            uIVistaElementosList.SelectedItemsAsString = this.ArchivoPrueba2Params.UIVistaElementosListSelectedItemsAsString;

            // Clic '&Abrir' botón
            Mouse.Click(uIAbrirButton, new Point(62, 13));
        }

        public virtual ArchivoPrueba2Params ArchivoPrueba2Params
        {
            get
            {
                if ((this.mArchivoPrueba2Params == null))
                {
                    this.mArchivoPrueba2Params = new ArchivoPrueba2Params();
                }
                return this.mArchivoPrueba2Params;
            }
        }

        private ArchivoPrueba2Params mArchivoPrueba2Params;

        /// <summary>
        /// AddFileSelfservice: use 'AddFileSelfserviceParams' para pasar parámetros a este método.
        /// </summary>
        public void AddFileSelfservice(string file, string Documentos)
        {
            #region Variable Declarations
            WinPane uIKACTUSSmartPeopleGooPane = this.UIKACTUSSmartPeopleGooWindow.UIKACTUSSmartPeopleGooPane;
            WinComboBox uINombreComboBox = this.UIAbrirWindow.UIItemWindow.UINombreComboBox;
            WinButton uIAbrirButton = this.UIAbrirWindow.UIAbrirWindow1.UIAbrirButton;
            #endregion

            //// Clic 'KACTUS Smart People - Google Chrome' panel
            //Mouse.Click(uIKACTUSSmartPeopleGooPane, new Point(374, 680));

            //// Seleccionar 'Archivo Prueba.pdf' en cuadro combinado 'Nombre:'
            //uINombreComboBox.EditableItem = this.AddFileSelfserviceParams.UINombreComboBoxEditableItem;

            //// Clic '&Abrir' botón
            //Mouse.Click(uIAbrirButton, new Point(43, 7));

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

        public virtual AddFileSelfserviceParams AddFileSelfserviceParams
        {
            get
            {
                if ((this.mAddFileSelfserviceParams == null))
                {
                    this.mAddFileSelfserviceParams = new AddFileSelfserviceParams();
                }
                return this.mAddFileSelfserviceParams;
            }
        }

        private AddFileSelfserviceParams mAddFileSelfserviceParams;

        /// <summary>
        /// AbrirSoporteSolicitud
        /// </summary>
        public void AbrirSoporteSolicitud()
        {
            #region Variable Declarations
            WinButton uIAbrirButton = this.UIAbrirWindow.UIAbrirWindow1.UIAbrirButton;
            #endregion

            // Clic '&Abrir' botón
            Mouse.Click(uIAbrirButton, new Point(56, 14));
        }

        /// <summary>
        /// DocumentosCesantias: use 'DocumentosCesantiasParams' para pasar parámetros a este método.
        /// </summary>
        public void DocumentosCesantias()
        {
            #region Variable Declarations
            WinPane uIKACTUSSmartPeopleGooPane = this.UIKACTUSSmartPeopleGooWindow.UIKACTUSSmartPeopleGooPane;
            WinComboBox uINombreComboBox = this.UIAbrirWindow.UIItemWindow.UINombreComboBox;
            WinButton uIAbrirButton = this.UIAbrirWindow.UIAbrirWindow1.UIAbrirButton;
            #endregion

            // Clic 'KACTUS Smart People - Google Chrome' panel
            Mouse.Click(uIKACTUSSmartPeopleGooPane, new Point(958, 640));

            // Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
            uINombreComboBox.EditableItem = this.DocumentosCesantiasParams.UINombreComboBoxEditableItem;

            // Clic '&Abrir' botón
            Mouse.Click(uIAbrirButton, new Point(24, 15));
        }

        public virtual DocumentosCesantiasParams DocumentosCesantiasParams
        {
            get
            {
                if ((this.mDocumentosCesantiasParams == null))
                {
                    this.mDocumentosCesantiasParams = new DocumentosCesantiasParams();
                }
                return this.mDocumentosCesantiasParams;
            }
        }

        private DocumentosCesantiasParams mDocumentosCesantiasParams;

        /// <summary>
        /// SelfserviceCesan: use 'SelfserviceCesanParams' para pasar parámetros a este método.
        /// </summary>
        public void SelfserviceCesan(string file, string Documentos)
        {
            #region Variable Declarations
            WinPane uIKACTUSSmartPeopleGooPane = this.UIKACTUSSmartPeopleGooWindow.UIKACTUSSmartPeopleGooPane;
            WinComboBox uINombreComboBox = this.UIAbrirWindow.UIItemWindow.UINombreComboBox;
            WinButton uIAbrirButton = this.UIAbrirWindow.UIAbrirWindow1.UIAbrirButton;
            #endregion

            //// Clic 'KACTUS Smart People - Google Chrome' panel
            //Mouse.Click(uIKACTUSSmartPeopleGooPane, new Point(1062, 643));

            //// Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
            //uINombreComboBox.EditableItem = this.SelfserviceCesanParams.UINombreComboBoxEditableItem;

            //// Clic '&Abrir' botón
            //Mouse.Click(uIAbrirButton, new Point(32, 18));

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

        public virtual SelfserviceCesanParams SelfserviceCesanParams
        {
            get
            {
                if ((this.mSelfserviceCesanParams == null))
                {
                    this.mSelfserviceCesanParams = new SelfserviceCesanParams();
                }
                return this.mSelfserviceCesanParams;
            }
        }

        private SelfserviceCesanParams mSelfserviceCesanParams;

        /// <summary>
        /// SelfserviceCesan1: use 'SelfserviceCesan1Params' para pasar parámetros a este método.
        /// </summary>
        public void SelfserviceCesan1(string file, string Documentos)
        {
            #region Variable Declarations
            WinPane uIKACTUSSmartPeopleGooPane = this.UIKACTUSSmartPeopleGooWindow.UIKACTUSSmartPeopleGooPane;
            WinComboBox uINombreComboBox = this.UIAbrirWindow.UIItemWindow.UINombreComboBox;
            WinButton uIAbrirButton = this.UIAbrirWindow.UIAbrirWindow1.UIAbrirButton;
            #endregion

            //// Clic 'KACTUS Smart People - Google Chrome' panel
            //Mouse.Click(uIKACTUSSmartPeopleGooPane, new Point(994, 690));

            //// Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
            //uINombreComboBox.EditableItem = this.SelfserviceCesan1Params.UINombreComboBoxEditableItem;

            //// Clic '&Abrir' botón
            //Mouse.Click(uIAbrirButton, new Point(38, 12));

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

        public virtual SelfserviceCesan1Params SelfserviceCesan1Params
        {
            get
            {
                if ((this.mSelfserviceCesan1Params == null))
                {
                    this.mSelfserviceCesan1Params = new SelfserviceCesan1Params();
                }
                return this.mSelfserviceCesan1Params;
            }
        }

        private SelfserviceCesan1Params mSelfserviceCesan1Params;

        /// <summary>
        /// AddSelfserviceReqpe: use 'AddSelfserviceReqpeParams' para pasar parámetros a este método.
        /// </summary>
        public void AddSelfserviceReqpe(string file, string Documentos)
        {
            #region Variable Declarations
            WinPane uIKACTUSSmartPeopleGooPane = this.UIKACTUSSmartPeopleGooWindow.UIKACTUSSmartPeopleGooPane;
            WinComboBox uINombreComboBox = this.UIAbrirWindow.UIItemWindow.UINombreComboBox;
            WinButton uIAbrirButton = this.UIAbrirWindow.UIAbrirWindow1.UIAbrirButton;
            #endregion

            //// Clic 'KACTUS Smart People - Google Chrome' panel
            //Mouse.Click(uIKACTUSSmartPeopleGooPane, new Point(405, 589));

            //// Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
            //uINombreComboBox.EditableItem = this.AddSelfserviceReqpeParams.UINombreComboBoxEditableItem;

            //// Clic '&Abrir' botón
            //Mouse.Click(uIAbrirButton, new Point(32, 13));

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

        public virtual AddSelfserviceReqpeParams AddSelfserviceReqpeParams
        {
            get
            {
                if ((this.mAddSelfserviceReqpeParams == null))
                {
                    this.mAddSelfserviceReqpeParams = new AddSelfserviceReqpeParams();
                }
                return this.mAddSelfserviceReqpeParams;
            }
        }

        private AddSelfserviceReqpeParams mAddSelfserviceReqpeParams;
    }
    /// <summary>
    /// Parámetros que se van a pasar a 'ArchivoPrueba2'
    /// </summary>
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class ArchivoPrueba2Params
    {

        #region Fields
        /// <summary>
        /// Seleccionar '' en cuadro de lista 'Vista Elementos'
        /// </summary>
        public string UIVistaElementosListSelectedItemsAsString = "";
        #endregion
    }
    /// <summary>
    /// Parámetros que se van a pasar a 'AddFileSelfservice'
    /// </summary>
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class AddFileSelfserviceParams
    {

        #region Fields
        /// <summary>
        /// Seleccionar 'Archivo Prueba.pdf' en cuadro combinado 'Nombre:'
        /// </summary>
        public string UINombreComboBoxEditableItem = "Archivo Prueba.pdf";
        #endregion
    }
    /// <summary>
    /// Parámetros que se van a pasar a 'DocumentosCesantias'
    /// </summary>
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class DocumentosCesantiasParams
    {

        #region Fields
        /// <summary>
        /// Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
        /// </summary>
        public string UINombreComboBoxEditableItem = "ArchivoPrueba.pdf";
        #endregion
    }
    /// <summary>
    /// Parámetros que se van a pasar a 'SelfserviceCesan'
    /// </summary>
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class SelfserviceCesanParams
    {

        #region Fields
        /// <summary>
        /// Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
        /// </summary>
        public string UINombreComboBoxEditableItem = "ArchivoPrueba.pdf";
        #endregion
    }
    /// <summary>
    /// Parámetros que se van a pasar a 'SelfserviceCesan1'
    /// </summary>
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class SelfserviceCesan1Params
    {

        #region Fields
        /// <summary>
        /// Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
        /// </summary>
        public string UINombreComboBoxEditableItem = "ArchivoPrueba.pdf";
        #endregion
    }
    /// <summary>
    /// Parámetros que se van a pasar a 'AddSelfserviceReqpe'
    /// </summary>
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class AddSelfserviceReqpeParams
    {

        #region Fields
        /// <summary>
        /// Seleccionar 'ArchivoPrueba.pdf' en cuadro combinado 'Nombre:'
        /// </summary>
        public string UINombreComboBoxEditableItem = "ArchivoPrueba.pdf";
        #endregion
    }
}
