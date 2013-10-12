using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Reflection;
using Microsoft.Office.Core;

namespace Word7AddIn
{
    public partial class ThisAddIn
    {
        _CommandBarButtonEvents_ClickEventHandler eventHandler;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {

                eventHandler = new _CommandBarButtonEvents_ClickEventHandler(MyButton_Click);
                Word.Application applicationObject = Globals.ThisAddIn.Application as Word.Application;
                applicationObject.WindowBeforeRightClick += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(App_WindowBeforeRightClick);
 
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error: " + exception.Message);
            }
        }

        void App_WindowBeforeRightClick(Microsoft.Office.Interop.Word.Selection Sel, ref bool Cancel)
        {
            try
            {
                this.AddItem();
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error: " + exception.Message);
            }
            
        }
        private void AddItem()
        {
            Word.Application applicationObject = Globals.ThisAddIn.Application as Word.Application;
            CommandBarButton commandBarButton = applicationObject.CommandBars.FindControl(MsoControlType.msoControlButton, missing, "HELLO_TAG", missing) as CommandBarButton;
            if (commandBarButton != null)
            {
                System.Diagnostics.Debug.WriteLine("Found button, attaching handler");
                commandBarButton.Click += eventHandler;
                return;
            }
            CommandBar popupCommandBar = applicationObject.CommandBars["Text"];
            bool isFound = false;
            foreach (object _object in popupCommandBar.Controls)
            {
                CommandBarButton _commandBarButton = _object as CommandBarButton;
                if (_commandBarButton == null) continue;
                if (_commandBarButton.Tag.Equals("HELLO_TAG"))
                {
                    isFound = true;
                    System.Diagnostics.Debug.WriteLine("Found existing button. Will attach a handler.");
                    commandBarButton.Click += eventHandler;
                    break;
                }
            }
            if (!isFound)
            {
                commandBarButton = (CommandBarButton)popupCommandBar.Controls.Add(MsoControlType.msoControlButton, missing, missing, missing, true);
                System.Diagnostics.Debug.WriteLine("Created new button, adding handler");
                commandBarButton.Click += eventHandler;
                commandBarButton.Caption = "Hello !!!";
                commandBarButton.FaceId = 356;
                commandBarButton.Tag = "HELLO_TAG";
                commandBarButton.BeginGroup = true;
            }
        }

        private void RemoveItem()
        {
            Word.Application applicationObject = Globals.ThisAddIn.Application as Word.Application;
            CommandBar popupCommandBar = applicationObject.CommandBars["Text"];
            foreach (object _object in popupCommandBar.Controls)
            {
                CommandBarButton commandBarButton = _object as CommandBarButton;
                if (commandBarButton == null) continue;
                if (commandBarButton.Tag.Equals("HELLO_TAG"))
                {
                    popupCommandBar.Reset();
                }
            }
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Word.Application App = Globals.ThisAddIn.Application as Word.Application;
            App.WindowBeforeRightClick -= new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(App_WindowBeforeRightClick);

        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion


        //Event Handler for the button click
        private void MyButton_Click(CommandBarButton cmdBarbutton, ref bool cancel)
        {
            System.Windows.Forms.MessageBox.Show("Hello !!! Happy Programming", "Hello !!!");
            Globals.ThisAddIn.Application.Selection.InsertAfter("I love CodeProject" + Environment.NewLine);
            Globals.ThisAddIn.Application.Selection.InsertAfter("Author: " + "Md. Marufuzzaman" + Environment.NewLine);
            Globals.ThisAddIn.Application.Selection.InsertAfter("Thanks To : " +  Environment.UserName + Environment.NewLine);
            Globals.ThisAddIn.Application.Selection.InsertAfter("Current time is :" +  DateTime.Now.ToLongTimeString() + Environment.NewLine);
            RemoveItem();

        }
    }
}
