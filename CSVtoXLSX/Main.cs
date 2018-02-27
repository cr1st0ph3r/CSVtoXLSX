using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace CSVtoXLSX
{
    public partial class Main : Form
    {
        private string Source = "";
        private bool Valid = false;
        private string log = AppDomain.CurrentDomain.BaseDirectory + @"Log.txt";
        private string Temp = AppDomain.CurrentDomain.BaseDirectory + @"temp.csv";
        private const int WIN_1252_CP = 1252;

        public Main()
        {
            InitializeComponent();
            txtDestino.Text = AppDomain.CurrentDomain.BaseDirectory + @"Output.xlsx";
        }

        //retrieve a file name on a drag and drop event
        protected bool GetFilename(out string filename, DragEventArgs e)
        {
            bool ret = false;
            filename = String.Empty;

            if ((e.AllowedEffect & DragDropEffects.Copy) == DragDropEffects.Copy)
            {
                Array data = ((IDataObject) e.Data).GetData("FileName") as Array;
                if (data != null)
                {
                    if ((data.Length == 1) && (data.GetValue(0) is String))
                    {
                        filename = ((string[]) data)[0];
                        string ext = Path.GetExtension(filename).ToLower();
                        if (ext == ".csv")
                        {
                            ret = true;
                        }
                    }
                }
            }

            return ret;
        }

        private void DoTheStuff()
        {
            //Set default variables
            string dCSV = Temp;
            string dXLSX = txtDestino.Text;
            string dLog = log;
            string dMerge = "no";
            object m = Type.Missing;

            try
            {

                    //Determine destination file existence
                    bool fileExists = File.Exists(dXLSX);

                    //Error checking
                    if (dMerge != "true" || fileExists == false)
                    {
                        dMerge = "false";
                    }

                    //Variables
                    Application excelApp = new Application();
                    bool merge = Boolean.Parse(dMerge);

                    //Open CSV
                    Workbook csv = excelApp.Workbooks.Open(dCSV);
                    Worksheet ws = csv.ActiveSheet;
                    csv.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;

                    //Get last row and column
                    int lastRow = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                    int lastCol = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

                    //Delete uneeded rows
                    ws.Rows[lastRow].Delete();
                    ws.Rows[lastRow - 1].Delete();
                    ws.Rows[2].Delete();

                    //Formatting
                    ws.Range[ws.Cells[1, 1], ws.Cells[1, lastCol]].Font.Bold = true;
                    ws.Range[ws.Cells[1, 1], ws.Cells[lastRow, lastCol]].Columns.AutoFit();

                    //Debugging
                    lbLog.Items.Add("Utima linha: " + lastRow);
                    lbLog.Items.Add("Utima Column: " + lastCol);
                    lbLog.Items.Add("juntar?: " + merge);

          
                //Merge or not, then save as XLSX
                if (fileExists == true && merge == true)
                    {
                        Workbook mergeXLSX = excelApp.Workbooks.Open(dXLSX);
                        ws.Move(Type.Missing, mergeXLSX.Worksheets[mergeXLSX.Worksheets.Count]);

                        mergeXLSX.Save();
                    }
                    else
                    {
                        excelApp.DisplayAlerts = false;
                        csv.SaveAs(dXLSX, XlFileFormat.xlOpenXMLWorkbook);
                      
                        excelApp.DisplayAlerts = true;
                    }

                    excelApp.Quit();
                
            }
            catch (Exception e)
            {
                //Create timestamp
                DateTime dt = DateTime.Now;
                string timestamp = dt.ToString("[yyyy/MM/dd HH:mm:ss]");

                //Create error string
                StringBuilder sb = new StringBuilder();
                sb.Append(timestamp + ": ");
                sb.Append(e.ToString());

                //Append to error log
                StreamWriter sw = new StreamWriter(dLog, true);
                sw.WriteLine("============================================================");
                sw.WriteLine("============= Args: CSV, XLSX, Log File, Merge =============");
                sw.WriteLine("============================================================");
                lbLog.Items.Add("============================================================");
                lbLog.Items.Add("============= Args: CSV, XLSX, Log File, Merge =============");
                lbLog.Items.Add("============================================================");
                lbLog.Items.Add(sb);
                sw.WriteLine(sb);
                sw.Close();
            }
        }

        //this is a temporary fix to put the file in the right encoding.
        //This is done by copying the file into a temporary file within the Executable
        //Directory and setting its encoding while saving, then using it to perform the conversion.
        //While being a sloppy solution it works well. nevertheless it need to be changed to something more suitable
        private void copySheet()
        {
            FileStream fs = new FileStream(Source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            StreamReader reader = new StreamReader(fs,Encoding.UTF8);
            var t = reader.ReadToEnd();
            System.IO.File.WriteAllText(Temp, t, Encoding.GetEncoding(WIN_1252_CP));
        

        }

        #region Events
        private void Main_DragDrop(object sender, DragEventArgs e)
        {
            Valid = GetFilename(out Source, e);
            if (Valid)
            {

                txtOrigem.Text = Source;
                lbLog.Items.Add("Item carregado com sucesso");
            }
            else
            {
                lbLog.Items.Add("Só aceito arquivos em CSV");
            }
        }
        private void Main_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }
        private void btnStart_Click(object sender, EventArgs e)
        {

            copySheet();
            DoTheStuff();

        }
        private void btnDefinirDestino_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    txtDestino.Text = fbd.SelectedPath + "Output.xlsx";
                }
            }

        }
        #endregion







    }
}
