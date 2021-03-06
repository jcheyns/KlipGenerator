﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Utility.Excel;
using Utility.String;
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

        /// <summary>
        /// Depending on the workbookname we will activat the addin.
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            RibbonKlip KlipRibbon = (RibbonKlip)Globals.Ribbons.GetRibbon(typeof(RibbonKlip));            
            Excel.Worksheet ws = Application.ActiveSheet;
            //Is there a pattern for this sheet?
            string pattern = GetPattern( Wb.NameWithoutExtension());
            if (!string.IsNullOrEmpty(pattern) && ws!=null && Regex.IsMatch(ws.Name, pattern))
            {
                KlipRibbon.EnableGeneration = true;
            }
            else
            {
                KlipRibbon.EnableGeneration = false;
            }
        }

        public string GetPattern(string aWorkbookName) {
        
            string patternKey=GetPatternKey( aWorkbookName);
            if (patternKey != null)
            {
                return ConfigurationManager.AppSettings.Get("ActiveNamePattern|" + patternKey);
            }
            return null;
        }

        public string GetPatternKey(string aWorkbookName)
        {
            string workbookToActivePatternKey = ConfigurationManager.AppSettings.Get("WorkbookToActivePatternKey");
            if (string.IsNullOrEmpty(workbookToActivePatternKey))
            {
                return null;
            };
            string patternKey = aWorkbookName;

            foreach (KeyValuePair<string, string> kvp in workbookToActivePatternKey.ToDict())
            {
                if (Regex.IsMatch(aWorkbookName, kvp.Key))
                {
                    return kvp.Value;                    
                }
            }
            return null;
        }

        private void EnableKlipRibbon() {

        }

        private void Application_SheetActivate(object Sh)
        {
            Excel.Worksheet ws = Sh as Excel.Worksheet;
            Excel.Workbook wb = ws.Parent as Excel.Workbook;
            Application_WorkbookActivate(wb);
            /*RibbonKlip KlipRibbon = (RibbonKlip)Globals.Ribbons.GetRibbon(typeof(RibbonKlip));
                       
            string pattern=ConfigurationManager.AppSettings.Get("ActiveNamePattern|" + wb.NameWithoutExtension());
            if (!string.IsNullOrEmpty(pattern) && Regex.IsMatch(ws.Name, pattern))
            {
                KlipRibbon.EnableGeneration = true;
            }
            else {
                KlipRibbon.EnableGeneration = false;
            }*/
        }

        private KlipWriter kw;

        protected override object RequestComAddInAutomationService()
        {
            if (kw == null)
                kw = new KlipWriter();

            return kw;
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
