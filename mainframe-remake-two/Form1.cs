using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace mainframe_remake_two
{
    public partial class Main_Form : Form
    {
        public Main_Form()
        {
            InitializeComponent();
        }

        private void BtnBrowseData_Click(object sender, EventArgs e)
        {
            DialogResult result = openFile.ShowDialog();
            if (result == DialogResult.OK)
            {
                string file = openFile.FileName;
                try
                {
                    txtShowData.Text = file;
                }
                catch (IOException)
                {

                }
            }

            if (txtShowData.Text.Length != 0 && txtShowLookup.Text.Length != 0)
            {
                btnRun.Enabled = true;
                lblProgress.Text = "Ready";
            } else if (txtShowData.Text.Length == 0 && txtShowLookup.Text.Length == 0)
            {
                lblProgress.Text = "Awaiting Data and Lookup...";

            }
            else if (txtShowData.Text.Length == 0)
            {
                lblProgress.Text = "Awaiting Data...";
            } else if (txtShowLookup.Text.Length == 0)
            {
                lblProgress.Text = "Awaiting Lookup...";

            }
            else
            {
                btnRun.Enabled = false;

            }
            

        }

        private void BtnBrowseLookup_Click(object sender, EventArgs e)
        {
            DialogResult result = openFile.ShowDialog();
            if (result == DialogResult.OK)
            {
                string file = openFile.FileName;
                try
                {
                    txtShowLookup.Text = file;
                }
                catch (IOException)
                {

                }
            }

            if (txtShowData.Text.Length != 0 && txtShowLookup.Text.Length != 0)
            {
                btnRun.Enabled = true;
                lblProgress.Text = "Ready";
            }
            else if (txtShowData.Text.Length == 0 && txtShowLookup.Text.Length == 0)
            {
                lblProgress.Text = "Awaiting Data and Lookup...";

            }
            else if (txtShowData.Text.Length == 0)
            {
                lblProgress.Text = "Awaiting Data...";
            }
            else if (txtShowLookup.Text.Length == 0)
            {
                lblProgress.Text = "Awaiting Lookup...";

            }
            else
            {
                btnRun.Enabled = false;

            }
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            if (progWorker.IsBusy != true)
            {
                DialogResult dialogResult = MessageBox.Show("Are you sure you want to run the report?\nEstimated run time: ~15 minutes", "Confirmation", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    progReport.Maximum = 100;
                    progReport.Value = 0;

                    btnBrowseData.Enabled = false;
                    btnBrowseLookup.Enabled = false;
                    btnRun.Enabled = false;

                    progWorker.RunWorkerAsync();
                }
                
            }
        }

        private void progWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var background = sender as BackgroundWorker;

            background.ReportProgress(0, new Tuple<string> ("Opening Excel App..."));

            string dataFilePath = txtShowData.Text;
            string lookupFilePath = txtShowLookup.Text;

            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Workbook lookupWB;

            try
            {
                oXL = new Excel.Application();
                oXL.Visible = false;

                background.ReportProgress(3, new Tuple<string>("Opening Workbooks..."));
                //Get existing workbook
                oWB = oXL.Workbooks.Open(dataFilePath, ReadOnly:false);
                lookupWB = oXL.Workbooks.Open(lookupFilePath, ReadOnly: false);

                background.ReportProgress(48, new Tuple<string>("Copying Lookup table..."));

                var xlSheets = oWB.Worksheets;
                var xlCharts = oWB.Charts;
                if (xlCharts.Count != 0)
                {
                    xlCharts.Delete();
                }
                
                lookupWB.Sheets[1].Name = "Lookup";
                int lastSheet = xlSheets.Count;
                lookupWB.Sheets[1].Copy(Type.Missing, xlSheets[lastSheet]);
                lookupWB.Close(0);



                background.ReportProgress(50, new Tuple<string>("Going through sheets..."));
                List<string> datasheetNames = new List<string> { "warton2 data", "warton4 data", "brought data", "chad data" };
                Excel.Worksheet PivotSheet;
                int percent = 50;
                foreach(Excel.Worksheet sheet in xlSheets)
                {
                    string name = sheet.Name.ToLower();
                    if (name != "warton2 data" && name != "warton4 data" && name != "brought data" && name != "chad data")
                    {
                        //xlSheets[sheet.Name].Delete();
                    } else
                    {
                        percent += 7;
                        background.ReportProgress(percent, new Tuple<string>("Converting job names to prefixes..."));
                        InsertTruncColumn(sheet);
                        background.ReportProgress(percent, new Tuple<string>("Inserting App column..."));
                        InsertAppColumn(sheet);
                        percent += 3;
                        background.ReportProgress(percent, new Tuple<string>("Creating New Pivot..."));
                        PivotSheet = xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
                        string pivotName = sheet.Name.Split(' ')[0] + " Pivot New";
                        PivotSheet.Name = pivotName;
                        CreatePivotSheet(oWB, sheet, PivotSheet, pivotName, false);


                    }
                }
                background.ReportProgress(95, new Tuple<string>("Formatting new data..."));
                CreateConsolidatedSheet(oXL, oWB, xlSheets);
                background.ReportProgress(100, new Tuple<string>("Opening Report..."));

                //oWB.SaveAs(@"C:\Users\Ross is the best\Desktop\MainframeTesting\newbook.xlsx");
                oWB.Close(0);
                oXL.Visible = true;
            } catch (Exception exception)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, exception.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, exception.Source);
                errorMessage = String.Concat(errorMessage, exception.InnerException);
                MessageBox.Show(errorMessage, "Error");
            }

            

        }

        private void progWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progReport.Value = e.ProgressPercentage;
            var args =(Tuple<string>) e.UserState;

            lblProgress.Text = args.Item1;
        }
        private void progWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                //Cancelled
                

            }
            else if (e.Error != null)
            {
                //Error
                

            }
            else
            {
                

            }
            btnBrowseData.Enabled = true;
            btnBrowseLookup.Enabled = true;
            btnRun.Enabled = true;
            lblProgress.Text = "Completed";
        }

        private void InsertTruncColumn(Excel.Worksheet oWS) {
            Excel.Range oRng;

            oRng = oWS.Columns["K"];

            oWS.Cells[1, 11] = "Truncated";

            var lastUsedRow = GetLastUsedRow(oWS);

            oRng = oWS.get_Range("K2", "K" + lastUsedRow);
            oRng.FormulaR1C1 = "=LEFT(RC[-8],3)";

            AlterEdgeCases(oWS);
            
        }

        private void AlterEdgeCases(Excel.Worksheet oWS)
        {
            var lastUsedRow = GetLastUsedRow(oWS);

            //
            // MCO
            //

            //Find edge cases
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;

            Excel.Range searchRange = oWS.get_Range("K2", "K" + lastUsedRow);
            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = searchRange.Find("MCO", Missing.Value,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
            Missing.Value, Missing.Value);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.FormulaR1C1 = "=LEFT(RC[-8],4)";

                currentFind = searchRange.FindNext(currentFind);
            }

            //
            // OMV
            //

            currentFind = null;
            firstFind = null;
            
            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = searchRange.Find("OMV", Missing.Value,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
            Missing.Value, Missing.Value);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.FormulaR1C1 = "=LEFT(RC[-8],4)";

                currentFind = searchRange.FindNext(currentFind);
            }

            //
            // XCF
            //

            currentFind = null;
            firstFind = null;

            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = searchRange.Find("XCF", Missing.Value,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
            Missing.Value, Missing.Value);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.FormulaR1C1 = "=LEFT(RC[-8],5)";

                currentFind = searchRange.FindNext(currentFind);
            }

            //
            // BSL
            //

            currentFind = null;
            firstFind = null;

            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = searchRange.Find("BSL", Missing.Value,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
            Missing.Value, Missing.Value);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.FormulaR1C1 = "=LEFT(RC[-8],4)";

                currentFind = searchRange.FindNext(currentFind);
            }

            //
            // BRE
            //

            currentFind = null;
            firstFind = null;

            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = searchRange.Find("BRE", Missing.Value,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
            Missing.Value, Missing.Value);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.FormulaR1C1 = "=LEFT(RC[-8],4)";

                currentFind = searchRange.FindNext(currentFind);
            }

            //
            // BDB
            //

            currentFind = null;
            firstFind = null;

            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = searchRange.Find("BDB", Missing.Value,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
            Missing.Value, Missing.Value);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.FormulaR1C1 = "=LEFT(RC[-8],6)";

                currentFind = searchRange.FindNext(currentFind);
            }
        }

        private void InsertAppColumn(Excel.Worksheet oWS)
        {
            Excel.Range oRng;

            oRng = oWS.Columns["D"];

            oRng.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            oWS.Cells[1, 4] = "APP";
            
            var lastUsedRow = GetLastUsedRow(oWS);

            oRng = oWS.get_Range("D2", "D" + lastUsedRow);
            oRng.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[8] & \"*\",Lookup!C[-3]:C[-2],2,FALSE), \"\")";
            
            
        }

        private int GetLastUsedRow(Excel._Worksheet oSheet)
        {
            return oSheet.Cells.Find("*", Missing.Value, Missing.Value, Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, Missing.Value, Missing.Value).Row;
        }

        private void CreatePivotSheet(Excel._Workbook workbook, Excel._Worksheet dataSheet, Excel._Worksheet pivotSheet, string tableName, bool unified)
        {
            //If consolidated sheet, range selects up until column C, if normal pivot, selects up until K
            string col;
            if (unified == true)
            {
                col = "C";
            }
            else
            {
                col = "K";
            }
            //Get last used row
            var lastUsedRow = GetLastUsedRow(dataSheet); //Select all data from starting cell to last column + row
            var dataRange = dataSheet.get_Range("A1", col + lastUsedRow);
            var pivotRange = pivotSheet.Cells[1, 1]; //Select target location
            var oPivotCache = (Excel.PivotCache)workbook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, dataRange); //Create cache specifying data is coming from a table
            var oPivotTable = (Excel.PivotTable)pivotSheet.PivotTables().Add(PivotCache: oPivotCache, TableDestination: pivotRange, TableName: tableName);//Create table
                       
            if (unified == true)//If consolidated sheet
            {
                //Set Row field to 'APP'
                var RowPivotField = (Excel.PivotField)oPivotTable.PivotFields("APP");
                RowPivotField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                //Set Values field to 'Total'
                var SumPivotField = (Excel.PivotField)oPivotTable.PivotFields("Total");
                SumPivotField.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                SumPivotField.Function = Excel.XlConsolidationFunction.xlSum;
                SumPivotField.Name = "CPU Time";
                //Set Column field to 'LPAR'
                var ColPivotField = (Excel.PivotField)oPivotTable.PivotFields("LPAR");
                ColPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;


            }

            else //If normal sheet
            {
                //Set Row field to 'APP'
                Excel.PivotField RowPivotField = (Excel.PivotField)oPivotTable.PivotFields("APP");
                RowPivotField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                //Set Values field to 'CPUTIME'
                Excel.PivotField SumPivotField = (Excel.PivotField)oPivotTable.PivotFields("CPUTIME");
                SumPivotField.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                SumPivotField.Function = Excel.XlConsolidationFunction.xlSum;
                SumPivotField.Name = "CPU Time";

            }


        }

        private void CreateConsolidatedSheet(Excel.Application oXL, Excel._Workbook workbook, Excel.Sheets xlSheets)
        {
            var consolidatedSheet = xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
            consolidatedSheet.Name = "Consolidated Data";

            var consolidatedPivot = xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
            consolidatedPivot.Name = "Consolidated Pivot";

            Excel.Worksheet pivotSheet;
            string name;
            int startRow = 0;
            int count = 0;
            int incTitle = 2;
            for (int i = 3; i <= 6; i++)
            {
                pivotSheet = (Excel.Worksheet)xlSheets[i];
                name = pivotSheet.Name.Split(' ')[0];
                CreateLPAR(pivotSheet, name);
                var lastRow = GetLastUsedRow(pivotSheet);
                var pRange = pivotSheet.get_Range("A" + incTitle, "C" + (lastRow - 1));
                Excel.Range cRange;
                
                if (count == 0)
                {
                    cRange = consolidatedSheet.Range["A" + (startRow + 1), "C" + (startRow + lastRow - 2)];
                    startRow += lastRow - 1;
                } else
                {
                    cRange = consolidatedSheet.Range["A" + (startRow), "C" + (startRow + lastRow - 2)];
                    startRow += lastRow - 3;
                }
                pRange.Copy(cRange);
                incTitle = 3;
                count = 1;
            }

            CreatePivotSheet(workbook, consolidatedSheet, consolidatedPivot, "Consolidated Pivot", true);
            var last = GetLastUsedRow(consolidatedPivot);
            var copyFrom = consolidatedPivot.Range["A2", "F" + last];

            var newBook = oXL.Workbooks.Add();
            var sheet = newBook.Sheets[1];
            sheet.Name = "CPUTime";
            var copyTo = sheet.Range["A1", "F" + last];
            copyFrom.Copy(copyTo);

            //
            //  Format data
            //
            last = GetLastUsedRow(sheet);
            var SourceRange = (Excel.Range)sheet.Range("A1", "F" + (last - 1));
            FormatAsTable(SourceRange, "CPUTime", "TableStyleMedium6");

            SourceRange = sheet.Range("A1", "F1");
            SourceRange.Interior.Color = Excel.XlRgbColor.rgbDarkCyan;
            SourceRange.Characters.Font.Size = 12;
            SourceRange.EntireColumn.AutoFit();
            SourceRange = sheet.Range("A" + last, "F" + last);
            SourceRange.Interior.Color = Excel.XlRgbColor.rgbDarkCyan;
            SourceRange.Characters.Font.Color = Color.White;
            SourceRange.Characters.Font.Bold = true;
            SourceRange.Characters.Font.Size = 12;


            // Remove (blank)

            var newLast = GetLastUsedRow(sheet);

            var blankRow = newLast - 1;

            var cellValue = (string)(sheet.Cells[blankRow, 1] as Excel.Range).Value;
            if (cellValue == "(blank)")
            {
                var tempRange = sheet.Range("A" + blankRow, "F" + blankRow);
                tempRange.EntireRow.Delete(Type.Missing);

            }

            //Update totals


            string[] collNames = { "Brought", "CHAD", "Warton2", "Warton4", "Grand Total" };
            var collNum = 0;

            for (int colls = 2; colls <= 6; colls++)
            {
                sheet.Cells[blankRow, colls].Formula = "=SUM(CPUTime[" + collNames[collNum] + "])";
                collNum++;
            }

        }

        private void CreateLPAR(Excel._Worksheet oSheet, string sheetName)
        {
            var lastUsedRow = GetLastUsedRow(oSheet);

            var oRng = oSheet.get_Range("C3", "C" + (lastUsedRow - 1));
            oRng.Value2 = sheetName;
            oSheet.Cells[2, 3] = "LPAR";           

        }
        private void FormatAsTable(Excel.Range SourceRange, string TableName, string TableStyleName)
        {
            SourceRange.Worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange,
            SourceRange, System.Type.Missing, XlYesNoGuess.xlYes, System.Type.Missing).Name =
                TableName;
            SourceRange.Select();
            SourceRange.Worksheet.ListObjects[TableName].TableStyle = TableStyleName;
        }

    }
}
