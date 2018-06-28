using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Utility.Excel;
using System.Configuration;
using System.Text.RegularExpressions;

namespace KlipGenerator
{
    public partial class KlipAddIn
    {



        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.SheetActivate += Application_SheetActivate;
            Application.WorkbookActivate += Application_WorkbookActivate;
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            RibbonKlip KlipRibbon = (RibbonKlip)Globals.Ribbons.GetRibbon(typeof(RibbonKlip));
            String name = Wb.NameWithoutExtension();
            Excel.Worksheet ws = Application.ActiveSheet;
            string pattern = ConfigurationManager.AppSettings.Get("ActiveNamePattern|" + name);
            if (!string.IsNullOrEmpty(pattern) && ws!=null && Regex.IsMatch(ws.Name, pattern))
            {
                KlipRibbon.EnableGeneration = true;
            }
            else
            {
                KlipRibbon.EnableGeneration = false;
            }
        }

        private void Application_SheetActivate(object Sh)
        {
            RibbonKlip KlipRibbon = (RibbonKlip)Globals.Ribbons.GetRibbon(typeof(RibbonKlip));
            Excel.Worksheet ws = Sh as Excel.Worksheet;
            Excel.Workbook wb=ws.Parent as Excel.Workbook;
            String name=wb.NameWithoutExtension();
            string pattern=ConfigurationManager.AppSettings.Get("ActiveNamePattern|" + name);
            if (!string.IsNullOrEmpty(pattern) && Regex.IsMatch(ws.Name, pattern))
            {
                KlipRibbon.EnableGeneration = true;
            }
            else {
                KlipRibbon.EnableGeneration = false;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
    }
}
