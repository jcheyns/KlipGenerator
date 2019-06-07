using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Utility.Excel;
using System.Configuration;

namespace KlipGenerator
{
    [ComVisible(true)]
    public interface IKlipWriter
    {
         void GenerateKlip(Excel.Worksheet aSheet, int minRow, int maxRow);
    }

    /// <summary>
    /// Calss to write Klip input file
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class KlipWriter : IKlipWriter
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="wbKlip">Workbook in which to write</param>
        /// <param name="matrix"> NxM matrix having N Coils to pass to Klip and M columns with attrbitues</param>
        /// <param name="hiddenrows">list of rows that were hiddedn in the sheet, they will be in the matrix, but need to skipped when writing.</param>
        /// <param name="colsToPrint">Column indexes in the matrix to print, in the order given in this list. when -1 is inthe lis the default will be used, see defaults</param>
        /// <param name="defaults">list of default values, will be used when colstoPrint is -1</param>
        /// <returns></returns>
        private Excel.Worksheet WriteKlipSheet(Excel.Workbook wbKlip, object[,] matrix, List<int> hiddenrows, List<int> colsToPrint, List<string> defaults)
        {
            Excel.Worksheet sh = wbKlip.Sheets["Klip Input"];
            object[,] result = (object[,])Array.CreateInstance(typeof(object), new int[] { matrix.GetUpperBound(0) - matrix.GetLowerBound(0) + 1 - hiddenrows.Count, colsToPrint.Count }, new int[] { 1, 1 });

            int minCol = result.GetLowerBound(1);
            int maxCol = result.GetUpperBound(1);

            int newI = result.GetLowerBound(0);
            for (int i = result.GetLowerBound(0); i <= result.GetUpperBound(0); i++)
            {
                if (!hiddenrows.Contains(i))
                {
                    for (int j = minCol; j <= maxCol; j++)
                    {
                        if (colsToPrint[j - minCol] == -1)
                        {
                            if (defaults[j - minCol] == "SEQ_NO")
                            {
                                result[newI, j] = newI;
                            }
                            else
                            {
                                result[newI, j] = defaults[j - minCol];
                            }
                        }
                        else
                        {
                            if (matrix[i, colsToPrint[j - minCol]] !=null && matrix[i, colsToPrint[j - minCol]].ToString().Contains(",")){
                                matrix[i, colsToPrint[j - minCol]] = matrix[i, colsToPrint[j - minCol]].ToString().Replace(",", ";");
                            }
                            result[newI, j] = matrix[i, colsToPrint[j - minCol]];
                        }
                    }
                    newI++;
                }
            }
            sh.Range[sh.Cells[2, 1], sh.Cells[matrix.GetUpperBound(0) - matrix.GetLowerBound(0) + 1 + 1, colsToPrint.Count]].Value2 = result;
            return sh;


        }

        /// <summary>
        /// The template to use in generating Klip File.
        /// </summary>
        /// <param name="wbName">Name of the workbook from which we want to create Klip input</param>
        /// <param name="shName">Name of the worksheet from which we want to create Klip input</param>
        /// <returns></returns>
        private string KlipTemplate(string wbName, string shName)
        {
            //based on the name of the workbook we get a key
            string patternKey = Globals.KlipAddIn.GetPatternKey(wbName);
            //using that key we find which template to use
            return ConfigurationManager.AppSettings.Get("Template|" + patternKey);
            //return "N:/Production Planning/HDGL Campaign Planning/HDGL Scheduling/Klip/Klip.xltm";
        }

        /// <summary>
        /// Generate a klip file from the data in the given sheet, taking into account rows minRow to maxRow
        /// </summary>
        /// <param name="aSheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        public void GenerateKlip(Excel.Worksheet aSheet, int minRow, int maxRow)
        {
            if (aSheet == null) {
                return;
            }
            string klipTemplate = KlipTemplate(aSheet.Parent.Name, aSheet.Name);
            Excel.Application app = Globals.KlipAddIn.Application;
            Excel.Worksheet sh = aSheet;
            Excel.Workbook wbKlip = app.Workbooks.Add(klipTemplate);
            sh.Activate();
            //the template has a colunm mapping, linking the 2 file formats
            if (wbKlip.Names.Item("Mapping").RefersToRange is Excel.Range mappingRg)
            {
                List<string> mapping = new List<string>();
                List<string> defaults = new List<string>();
                for (int r = mappingRg.Row + 1; r < mappingRg.Row + mappingRg.Rows.Count; r++)
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
                List<int> hiddenrows = new List<int>();
                for (int i = minRow; i <= maxRow; i++)
                {
                    if (sh.Rows[i].Hidden)
                    {
                        hiddenrows.Add(i - minRow + matrix.GetLowerBound(0));
                    }
                }
                Excel.Worksheet KlipSh = WriteKlipSheet(wbKlip, matrix, hiddenrows, colsToPrint, defaults);
                
                //do some formatting to make sure it matches what Klip is expecting
                Dictionary<string, int> klipColDict = KlipSh.ColumnDictionary();
                if (klipColDict.TryGetValue("MES_MFCT_ORDER_ID", out col))
                {
                    KlipSh.Columns[col].EntireColumn.NumberFormat = "0";
                }
                if (klipColDict.TryGetValue("MFCT_ORDER_NO", out col))
                {
                    KlipSh.Columns[col].EntireColumn.NumberFormat = "0";
                }
                if (klipColDict.TryGetValue("DUE_DATE",out col)) {
                    KlipSh.Columns[col].EntireColumn.NumberFormat = "m/d/yyyy";
                }
                
                string tms = DateTime.Now.ToString("yyyyMMddHHmmss");
                KlipSh.ExportToCSV("C:\\KlipFiles", string.Format("KlipIn_MDP{0}.csv", tms));
                wbKlip.Close(false);
            }
        }
    }
}
