using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;
using Excel=Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Utility.Excel;

namespace KlipGenerator
{
    public partial class RibbonKlip
    {
        private bool _EnableGeneration;
        public bool EnableGeneration {
            get { return _EnableGeneration; }
            set {
                _EnableGeneration = value;
                SetGenerateButtonsEnabled(value);
            }
        }

        private void SetGenerateButtonsEnabled(bool value)
        {
            btnGenerate.Enabled = value;
            btnGenerateFromSelection.Enabled = value;
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BtnGenerate_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sh = Globals.KlipAddIn.Application.ActiveSheet;
            KlipWriter kw = new KlipWriter();
            kw.GenerateKlip(sh,2,sh.FirstEmptyRow()-1);
        }


        private void BtnGenerateFromSelection_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range SelectionRg=Globals.KlipAddIn.Application.Selection;
            KlipWriter kw = new KlipWriter();
            kw.GenerateKlip(SelectionRg.Parent,SelectionRg.Row, SelectionRg.Row+SelectionRg.Rows.Count-1);
        }
    }
}
