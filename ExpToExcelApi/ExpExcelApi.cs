using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;

namespace ExpToExcelApi
{
    public class ExpExcelApi
    {
        /// <summary>
        /// test the excel exporting first before creating an implementation
        /// </summary>
        /// <param name="args"></param>
        //static void Main(string[] args)
        //{
        //    Console.WriteLine("exporting to excel. . . . . . . ."); 
        //    List<string> myfname = new List<string>();
        //    myfname.Add("Sonny");
        //    myfname.Add("Sonny");
        //    myfname.Add("Sonny");
        //    List<string> mylname = new List<string>();
        //    mylname.Add("Recio");
        //    mylname.Add("Recio");
        //    mylname.Add("Recio");
        //    mylname.Add("Recio");
        //    ExcelExport(myfname, 1);
        //    ExcelExport(mylname, 2);
        //}

        private Microsoft.Office.Interop.Excel.Application myapp = new Microsoft.Office.Interop.Excel.Application { Visible = false }; //{ Visible = true };
        private Workbook workbook;
        private Worksheet worksheet;

        /// <summary>
        /// use this to reset excel and instantiate it.
        /// </summary>
        public void ResetExcel()
        {
            myapp = new Microsoft.Office.Interop.Excel.Application(); // { Visible = true };
            isInitialized = false;
        }

        private bool ContinueSave { get; set; }
        private string SaveFilepath()
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = "Crew Monitoring";
            saveFileDialog.DefaultExt = ".xls";
            saveFileDialog.Filter = "Excel File |*.xls|All Files| *.*";
            saveFileDialog.Title = "Where to save Crew Monitoring?";
            switch (saveFileDialog.ShowDialog())
            {
                case DialogResult.OK:
                    switch (saveFileDialog.OverwritePrompt) //if there are existing excel file.
                    {
                        case true:
                            File.Delete(saveFileDialog.FileName); //delete the file to avoid overriding with Excel prompt. Else it will yield an error if the user tries to cancel the operation.
                            break;
                    }
                    ContinueSave = true;
                    break;
                case DialogResult.Cancel:
                    ContinueSave = false;
                    break;
                default:
                    ContinueSave = false;
                    break;
            }
            return saveFileDialog.FileName;
        }

        /// <summary>
        /// call this method to export and save to an excel file
        /// </summary>
        public void SaveExcel()
        {
            string path = SaveFilepath();
            if (ContinueSave.Equals(true))
            {
                workbook.SaveAs(Filename: path, FileFormat: XlFileFormat.xlWorkbookNormal);
                myapp.Application.Quit();
                Process.Start(path); //opens file
            }
            //return saveFileDialog.FileName;
            //return myapp.GetSaveAsFilename(FileFilter: "*.xls");
        }

        /// <summary>
        /// use this method if you're gonna add another Worksheet within an MS Excel Application.
        /// </summary>
        public void NewSheet()
        {
            //isInitialized = false;
            myapp.Worksheets.Add();
            worksheet = myapp.ActiveSheet;
        }

        private bool isInitialized = false;
        private void InitializeExcel()
        {
            workbook = myapp.Workbooks.Add();
            worksheet = myapp.ActiveSheet;
        }

        public void ExcelExport(List<string> value, int Column, string worksheetName = "Default")
        {
            /*
           * note: must be using .net framework 4 to make this simple excel code possible.
           */
            if (isInitialized == false)
            {
                InitializeExcel();
                isInitialized = true;
            }
            int row = 1; //use row variable to control each row and loop it through
            worksheet.Name = worksheetName; //"List of persons"; //sets the name of worksheet
            //worksheet.Columns.Value = "test";
            foreach (var item in value)
            {
                worksheet.Cells[row, Column].Value = item;
                row++;
            }
            worksheet.Rows[1].Font.Bold = true;
            //worksheet.Rows.AutoFilter(worksheet.Range["A1", "C1"]);
            worksheet.Columns.AutoFit(); //apply autofit to the columns so that all data would be seen without truncating or hiding excessive data.
            //worksheet.Rows.AutoFit();
        }
    }
}
