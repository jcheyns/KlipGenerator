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
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BtnGenerate_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sh = Globals.KlipAddIn.Application.ActiveSheet;
            Generate(2,sh.FirstEmptyRow()-1);
        }

        private void Generate(int minRow, int maxRow)
        {
            string klipTemplate = KlipTemplate("","");
            Excel.Application app =Globals.KlipAddIn.Application;
            Excel.Worksheet sh = app.ActiveSheet;
            Excel.Workbook wbKlip =app.Workbooks.Add(klipTemplate);
            sh.Activate();
            if (wbKlip.Names.Item("Mapping").RefersToRange is Excel.Range mappingRg)
            {
                List<string> mapping = new List<string>();
                List<string> defaults = new List<string>();
                for (int r = mappingRg.Row+1; r < mappingRg.Row + mappingRg.Rows.Count; r++)
                {
                    mapping.Add(mappingRg.Cells[r, 2].Text);
                    defaults.Add(mappingRg.Cells[r, 3].Text);
                }
                
                Dictionary<string, int> colDict = sh.ColumnDictionary();
                List<int> colsToPrint = new List<int>(mapping.Count);
                int col;
                for (int r = 0; r < mapping.Count; r++)
                {
                    if (colDict.TryGetValue(mapping[r], out col))
                    {
                        colsToPrint.Add(col);
                    }
                    else
                    {
                        colsToPrint.Add(-1);
                    }
                }
                Excel.Range data = sh.Range[sh.Cells[minRow, 1], sh.Cells[maxRow, colDict.Count]];
                object[,] matrix = data.Value2;
                Excel.Worksheet KlipSh=WriteKlipSheet(wbKlip,matrix,colsToPrint,defaults);
                //some hard coded stuff, to be replaced with acutal data
                Dictionary<string, int> klipColDict = KlipSh.ColumnDictionary();
                if (klipColDict.TryGetValue("MES_MFCT_ORDER_ID", out col))
                {
                    sh.Columns[col].EntireColumn.NumberFormat = "0";
                }
                if (klipColDict.TryGetValue("MFCT_ORDER_NO", out col))
                {
                    sh.Columns[col].EntireColumn.NumberFormat = "0";
                }
                //coat min max
                Dictionary<int, Tuple<int, int>> Coating=new Dictionary<int, Tuple<int, int>>();
                Coating[30]=new Tuple<int, int>(20,40);
                Coating[37] = new Tuple<int, int>(30, 50);
                Coating[40] = new Tuple<int, int>(38, 70);
                Coating[42] = new Tuple<int, int>(38, 70);
                Coating[45] = new Tuple<int, int>(40, 54);
                Coating[50] = new Tuple<int, int>(40, 60);
                Coating[51] = new Tuple<int, int>(46, 65);
                Coating[53] = new Tuple<int, int>(48, 65);
                Coating[55] = new Tuple<int, int>(50, 70);
                Coating[65] = new Tuple<int, int>(60, 90);
                Coating[66] = new Tuple<int, int>(61, 95);
                Coating[70] = new Tuple<int, int>(61, 95);
                Coating[73] = new Tuple<int, int>(65, 95);
                Coating[75] = new Tuple<int, int>(70, 100);
                Coating[80] = new Tuple<int, int>(75, 100);
                Coating[102] = new Tuple<int, int>(92, 145);
                Coating[105] = new Tuple<int, int>(96, 145);
                Coating[145] = new Tuple<int, int>(138, 160);
                Coating[148] = new Tuple<int, int>(142, 160);
                int zn;
                int znMin;
                int znMax;
                if (klipColDict.TryGetValue("ZL_BOTTOM", out zn) && klipColDict.TryGetValue("ZL_BOTTOM_MIN", out znMin) && klipColDict.TryGetValue("ZL_BOTTOM_MAX", out znMax)) {
                    for (int i = 2;  i < KlipSh.FirstEmptyRow(9); i++){
                        try
                        {
                            int coatingWt = Convert.ToInt16(KlipSh.Cells[i, zn].Value2);
                            if (Coating.TryGetValue(coatingWt, out Tuple<int, int> tol))
                            {
                                KlipSh.Cells[i, znMin] = tol.Item1;
                                KlipSh.Cells[i, znMax] = tol.Item2;
                            }
                        }
                        catch (Exception)
                        {
                            //just ignore
                        }
                        
                    }
                }
                if (klipColDict.TryGetValue("ZL_TOP", out zn) && klipColDict.TryGetValue("ZL_TOP_MIN", out znMin) && klipColDict.TryGetValue("ZL_TOP_MAX", out znMax))
                {
                    for (int i = 2; i < KlipSh.FirstEmptyRow(9); i++)
                    {
                        try
                        {
                            int coatingWt = Convert.ToInt16(KlipSh.Cells[i, zn].Value2);
                            if (Coating.TryGetValue(coatingWt, out Tuple<int, int> tol))
                            {
                                KlipSh.Cells[i, znMin] = tol.Item1;
                                KlipSh.Cells[i, znMax] = tol.Item2;
                            }
                        }
                        catch (Exception)
                        {
                            //just ignore
                        }
                    }
                }

                KlipSh.ExportToCSV("C:\\Test\\KlipInvTEST.csv");
                wbKlip.Close(false);
            }
        }

        private Excel.Worksheet WriteKlipSheet(Excel.Workbook wbKlip, object[,] matrix, List<int> colsToPrint, List<string> defaults)
        {
            Excel.Worksheet sh = wbKlip.Sheets["Klip Input"];
            object[,] result = (object[,])Array.CreateInstance(typeof(object), new int[] { matrix.GetUpperBound(0) - matrix.GetLowerBound(0) + 1, colsToPrint.Count }, new int[]{ 1, 1});            

            for (int i = result.GetLowerBound(0); i <= result.GetUpperBound(0); i++) {
                for (int j = result.GetLowerBound(1); j <= result.GetUpperBound(1); j++)
                {
                    if (colsToPrint[j- matrix.GetLowerBound(1)] == -1)
                    {
                        result[i, j] = defaults[j - matrix.GetLowerBound(1)];
                    }
                    else {
                        result[i, j] = matrix[i, colsToPrint[j- matrix.GetLowerBound(1)]];
                    }
                }
            }
            sh.Range[sh.Cells[2,1],sh.Cells[matrix.GetUpperBound(0) - matrix.GetLowerBound(0) + 1+1, colsToPrint.Count]].Value2 = result;
            return sh;

            
        }

        private string KlipTemplate(string wbName, string shName)
        {
            return "N:/Production Planning/HDGL Campaign Planning/HDGL Scheduling/Klip/Klip.xltm";
        }

        private void BtnGenerateFromSelection_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range SelectionRg=Globals.KlipAddIn.Application.Selection;

            Generate(SelectionRg.Row, SelectionRg.Row+SelectionRg.Rows.Count-1);
        }
    }
}
