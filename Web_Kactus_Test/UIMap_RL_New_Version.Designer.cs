﻿// ------------------------------------------------------------------------------
//  <auto-generated>
//      Este código lo generó el generador de pruebas automatizadas de IU.
//      Versión: 15.0.0.0
//
//      Los cambios realizados en este archivo pueden provocar un comportamiento incorrecto y se perderán si
//      se vuelve a generar el código.
//  </auto-generated>
// ------------------------------------------------------------------------------

namespace Web_Kactus_Test.UIMap_RL_New_VersionClasses
{
    using System;
    using System.CodeDom.Compiler;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Text.RegularExpressions;
    using System.Windows.Input;
    using Microsoft.VisualStudio.TestTools.UITest.Extension;
    using Microsoft.VisualStudio.TestTools.UITesting;
    using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
    using Mouse = Microsoft.VisualStudio.TestTools.UITesting.Mouse;
    using MouseButtons = System.Windows.Forms.MouseButtons;
    
    
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public partial class UIMap_RL_New_Version
    {
        
        /// <summary>
        /// Basura
        /// </summary>
        public void Basura()
        {
            #region Variable Declarations
            WinControl uIChromeLegacyWindowDocument = this.UIKactusRLReclutamientWindow.UIChromeLegacyWindowWindow.UIChromeLegacyWindowDocument;
            #endregion

            // Clic 'Chrome Legacy Window' documento
            Mouse.Click(uIChromeLegacyWindowDocument, new Point(791, 639));
        }
        
        #region Properties
        public UIKactusRLReclutamientWindow UIKactusRLReclutamientWindow
        {
            get
            {
                if ((this.mUIKactusRLReclutamientWindow == null))
                {
                    this.mUIKactusRLReclutamientWindow = new UIKactusRLReclutamientWindow();
                }
                return this.mUIKactusRLReclutamientWindow;
            }
        }
        
        public UIAbrirWindow UIAbrirWindow
        {
            get
            {
                if ((this.mUIAbrirWindow == null))
                {
                    this.mUIAbrirWindow = new UIAbrirWindow();
                }
                return this.mUIAbrirWindow;
            }
        }
        #endregion
        
        #region Fields
        private UIKactusRLReclutamientWindow mUIKactusRLReclutamientWindow;
        
        private UIAbrirWindow mUIAbrirWindow;
        #endregion
    }
    
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class UIKactusRLReclutamientWindow : WinWindow
    {
        
        public UIKactusRLReclutamientWindow()
        {
            #region Criterio de búsqueda
            this.SearchProperties[WinWindow.PropertyNames.Name] = ".:. Kactus RL Reclutamiento Web .:. - Google Chrome";
            this.SearchProperties[WinWindow.PropertyNames.ClassName] = "Chrome_WidgetWin_1";
            this.WindowTitles.Add(".:. Kactus RL Reclutamiento Web .:. - Google Chrome");
            #endregion
        }
        
        #region Properties
        public UIChromeLegacyWindowWindow UIChromeLegacyWindowWindow
        {
            get
            {
                if ((this.mUIChromeLegacyWindowWindow == null))
                {
                    this.mUIChromeLegacyWindowWindow = new UIChromeLegacyWindowWindow(this);
                }
                return this.mUIChromeLegacyWindowWindow;
            }
        }
        #endregion
        
        #region Fields
        private UIChromeLegacyWindowWindow mUIChromeLegacyWindowWindow;
        #endregion
    }
    
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class UIChromeLegacyWindowWindow : WinWindow
    {
        
        public UIChromeLegacyWindowWindow(UITestControl searchLimitContainer) : 
                base(searchLimitContainer)
        {
            #region Criterio de búsqueda
            this.SearchProperties[WinWindow.PropertyNames.ControlId] = "1065493824";
            this.WindowTitles.Add(".:. Kactus RL Reclutamiento Web .:. - Google Chrome");
            #endregion
        }
        
        #region Properties
        public WinControl UIChromeLegacyWindowDocument
        {
            get
            {
                if ((this.mUIChromeLegacyWindowDocument == null))
                {
                    this.mUIChromeLegacyWindowDocument = new WinControl(this);
                    #region Criterio de búsqueda
                    this.mUIChromeLegacyWindowDocument.SearchProperties[UITestControl.PropertyNames.ControlType] = "Document";
                    this.mUIChromeLegacyWindowDocument.WindowTitles.Add(".:. Kactus RL Reclutamiento Web .:. - Google Chrome");
                    #endregion
                }
                return this.mUIChromeLegacyWindowDocument;
            }
        }
        #endregion
        
        #region Fields
        private WinControl mUIChromeLegacyWindowDocument;
        #endregion
    }
    
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class UIAbrirWindow : WinWindow
    {
        
        public UIAbrirWindow()
        {
            #region Criterio de búsqueda
            this.SearchProperties[WinWindow.PropertyNames.Name] = "Abrir";
            this.SearchProperties[WinWindow.PropertyNames.ClassName] = "#32770";
            this.WindowTitles.Add("Abrir");
            #endregion
        }
        
        #region Properties
        public UIItemWindow UIItemWindow
        {
            get
            {
                if ((this.mUIItemWindow == null))
                {
                    this.mUIItemWindow = new UIItemWindow(this);
                }
                return this.mUIItemWindow;
            }
        }
        
        public UIAbrirWindow1 UIAbrirWindow1
        {
            get
            {
                if ((this.mUIAbrirWindow1 == null))
                {
                    this.mUIAbrirWindow1 = new UIAbrirWindow1(this);
                }
                return this.mUIAbrirWindow1;
            }
        }
        #endregion
        
        #region Fields
        private UIItemWindow mUIItemWindow;
        
        private UIAbrirWindow1 mUIAbrirWindow1;
        #endregion
    }
    
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class UIItemWindow : WinWindow
    {
        
        public UIItemWindow(UITestControl searchLimitContainer) : 
                base(searchLimitContainer)
        {
            #region Criterio de búsqueda
            this.SearchProperties[WinWindow.PropertyNames.ControlId] = "1148";
            this.SearchProperties[WinWindow.PropertyNames.Instance] = "2";
            this.WindowTitles.Add("Abrir");
            #endregion
        }
        
        #region Properties
        public WinComboBox UINombreComboBox
        {
            get
            {
                if ((this.mUINombreComboBox == null))
                {
                    this.mUINombreComboBox = new WinComboBox(this);
                    #region Criterio de búsqueda
                    this.mUINombreComboBox.SearchProperties[WinComboBox.PropertyNames.Name] = "Nombre:";
                    this.mUINombreComboBox.WindowTitles.Add("Abrir");
                    #endregion
                }
                return this.mUINombreComboBox;
            }
        }
        #endregion
        
        #region Fields
        private WinComboBox mUINombreComboBox;
        #endregion
    }
    
    [GeneratedCode("Generador de pruebas de UI codificadas", "15.0.26208.0")]
    public class UIAbrirWindow1 : WinWindow
    {
        
        public UIAbrirWindow1(UITestControl searchLimitContainer) : 
                base(searchLimitContainer)
        {
            #region Criterio de búsqueda
            this.SearchProperties[WinWindow.PropertyNames.ControlId] = "1";
            this.WindowTitles.Add("Abrir");
            #endregion
        }
        
        #region Properties
        public WinButton UIAbrirButton
        {
            get
            {
                if ((this.mUIAbrirButton == null))
                {
                    this.mUIAbrirButton = new WinButton(this);
                    #region Criterio de búsqueda
                    this.mUIAbrirButton.SearchProperties[WinButton.PropertyNames.Name] = "Abrir";
                    this.mUIAbrirButton.WindowTitles.Add("Abrir");
                    #endregion
                }
                return this.mUIAbrirButton;
            }
        }
        #endregion
        
        #region Fields
        private WinButton mUIAbrirButton;
        #endregion
    }
}
