using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddInSample
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }
        //private void AddButtonsToMenu()
        //{
        //    RibbonButton tempButton = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
        //    tempButton.Label = "Button 1";
        //    tempButton.ControlSize =
        //        Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
        //    tempButton.Description = "My Ribbon Button";
        //    tempButton.ShowImage = true;
        //    tempButton.Image = Properties.Resources.Горы;
        //    tempButton.KeyTip = "A1";
        //    Globals.Ribbons.GetRibbon(Ribbon).Add(tempButton);

        //}
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
