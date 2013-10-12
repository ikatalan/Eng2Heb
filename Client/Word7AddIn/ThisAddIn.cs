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
        _CommandBarButtonEvents_ClickEventHandler eng2HebEventHandler;

        private Dictionary<Char, Char> eng2heb;
        private Dictionary<Char, Char> heb2eng;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                eventHandler = new _CommandBarButtonEvents_ClickEventHandler(MyButton_Click);
                eng2HebEventHandler = new _CommandBarButtonEvents_ClickEventHandler(Eng2Heb_Click);

                Word.Application applicationObject = Globals.ThisAddIn.Application as Word.Application;
                applicationObject.WindowBeforeRightClick += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(App_WindowBeforeRightClick);
                
                eng2heb = new Dictionary<char,char>();
                eng2heb['q'] = '/';
                eng2heb['w'] = '\'';
                eng2heb['e'] = 'ק';
                eng2heb['r'] = 'ר';
                eng2heb['t'] = 'א';
                eng2heb['y'] = 'ט';
                eng2heb['u'] = 'ו';
                eng2heb['i'] = 'ן';
                eng2heb['o'] = 'ם';
                eng2heb['p'] = 'פ';
                eng2heb['['] = ']';
                eng2heb[']'] = '[';
                eng2heb['a'] = 'ש';
                eng2heb['s'] = 'ד';
                eng2heb['d'] = 'ג';
                eng2heb['f'] = 'כ';
                eng2heb['g'] = 'ע';
                eng2heb['h'] = 'י';
                eng2heb['j'] = 'ח';
                eng2heb['k'] = 'ל';
                eng2heb['l'] = 'ך';
                eng2heb[';'] = 'ף';
                eng2heb['\''] = ',';
                eng2heb['z'] = 'ז';
                eng2heb['x'] = 'ס';
                eng2heb['c'] = 'ב';
                eng2heb['v'] = 'ה';
                eng2heb['b'] = 'נ';
                eng2heb['n'] = 'מ';
                eng2heb['m'] = 'צ';
                eng2heb[','] = 'ת';
                eng2heb['.'] = 'ץ';
                eng2heb['/'] = '.';
                eng2heb['Q'] = '/';
                eng2heb['W'] = '\'';
                eng2heb['E'] = 'ק';
                eng2heb['R'] = 'ר';
                eng2heb['T'] = 'א';
                eng2heb['Y'] = 'ט';
                eng2heb['U'] = 'ו';
                eng2heb['I'] = 'ן';
                eng2heb['O'] = 'ם';
                eng2heb['P'] = 'פ';
                eng2heb['}'] = ']';
                eng2heb['}'] = '[';
                eng2heb['A'] = 'ש';
                eng2heb['S'] = 'ד';
                eng2heb['D'] = 'ג';
                eng2heb['F'] = 'כ';
                eng2heb['G'] = 'ע';
                eng2heb['H'] = 'י';
                eng2heb['J'] = 'ח';
                eng2heb['K'] = 'ל';
                eng2heb['L'] = 'ך';
                eng2heb[':'] = 'ף';
                eng2heb['\"'] = ',';
                eng2heb['Z'] = 'ז';
                eng2heb['X'] = 'ס';
                eng2heb['C'] = 'ב';
                eng2heb['V'] = 'ה';
                eng2heb['B'] = 'נ';
                eng2heb['N'] = 'מ';
                eng2heb['M'] = 'צ';
                eng2heb['>'] = 'ת';
                eng2heb['<'] = 'ץ';
                eng2heb['?'] = '.';

                

                heb2eng = new Dictionary<char, char>();
                heb2eng['/'] = 'q';
                heb2eng['\''] = 'w';
                heb2eng['ק'] = 'e';
                heb2eng['ר'] = 'r';
                heb2eng['א'] = 't';
                heb2eng['ט'] = 'y';
                heb2eng['ו'] = 'u';
                heb2eng['ן'] = 'i';
                heb2eng['ם'] = 'o';
                heb2eng['פ'] = 'p';
                heb2eng[']'] = '[';
                heb2eng['['] = ']';
                heb2eng['ש'] = 'a';
                heb2eng['ד'] = 's';
                heb2eng['ג'] = 'd';
                heb2eng['כ'] = 'f';
                heb2eng['ע'] = 'g';
                heb2eng['י'] = 'h';
                heb2eng['ח'] = 'j';
                heb2eng['ל'] = 'k';
                heb2eng['ך'] = 'l';
                heb2eng['ף'] = ';';
                heb2eng[','] = '\'';
                heb2eng['ז'] = 'z';
                heb2eng['ס'] = 'x';
                heb2eng['ב'] = 'c';
                heb2eng['ה'] = 'v';
                heb2eng['נ'] = 'b';
                heb2eng['מ'] = 'n';
                heb2eng['צ'] = 'm';
                heb2eng['ת'] = ',';
                heb2eng['ץ'] = '.';
                heb2eng['.'] = '/';
                
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
                commandBarButton.Caption = "Heb to Eng";
                commandBarButton.FaceId = 356;
                commandBarButton.Tag = "HELLO_TAG";
                commandBarButton.BeginGroup = true;


                commandBarButton = (CommandBarButton)popupCommandBar.Controls.Add(MsoControlType.msoControlButton, missing, missing, missing, true);
                System.Diagnostics.Debug.WriteLine("Created new button eng2heb, adding handler");
                commandBarButton.Click += eng2HebEventHandler;
                commandBarButton.Caption = "Eng to Heb";
                commandBarButton.FaceId = 356;
                commandBarButton.Tag = "ENG2HEB_TAG";
                commandBarButton.BeginGroup = false;


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
            string opposite = "";
            foreach (char chr in Globals.ThisAddIn.Application.Selection.Text.ToCharArray())
            {
                if (heb2eng.ContainsKey(chr))
                {
                    opposite += heb2eng[chr];
                }
                else
                {
                    opposite += chr;
                }
            }

            Globals.ThisAddIn.Application.Selection.Delete();
            Globals.ThisAddIn.Application.Selection.InsertAfter(opposite);
            RemoveItem();
        }

        private void Eng2Heb_Click(CommandBarButton cmdBarbutton, ref bool cancel)
        {
            string opposite = "";
            foreach (char chr in Globals.ThisAddIn.Application.Selection.Text.ToCharArray())
            {
                if (eng2heb.ContainsKey(chr))
                {
                    opposite += eng2heb[chr];
                }
                else
                {
                    opposite += chr;
                }
            }

            Globals.ThisAddIn.Application.Selection.Delete();
            Globals.ThisAddIn.Application.Selection.InsertAfter(opposite);
            RemoveItem();
        }

    }
}
