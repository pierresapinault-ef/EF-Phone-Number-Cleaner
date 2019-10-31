using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.ComponentModel;
using Microsoft.Win32;
using System.Windows;
using System.Windows.Media;
using System.Data;
using System.Xaml;
using System.Runtime.InteropServices;
using System.Deployment.Application;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Controls;

namespace PhoneNumberCleaner
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            mainWindow.Title = "Phone Number Cleaner " + getRunningVersion();
        }

        private string uploadedFilePath;
        private bool messageBird;
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private string Locale;
        private DataTable dt = null;
        private string newFileName;
        private int lastRow = 0;
        private Excel.Range range = null;
        public static string customerID = "customer_id";
        public static string MobilePhone = "mobilephone";
        public static string FirstName = "firstname";
        public static string LastName = "lastname";
        public static string customer_ID = "customer id";
        public static string Mobile_Phone = "mobile phone";
        public static string MobileNumber = "mobilenumber";
        public static string First_Name = "first name";
        public static string Last_Name = "last name";
        public static string SubscriberKey = "subscriberkey";
        public static string[] arr = { SubscriberKey, customerID, customer_ID, Mobile_Phone, MobilePhone, MobileNumber, FirstName, First_Name, LastName, Last_Name };
        public static string[] markets = { "JP", "KR", "TH", "TW", "VN", "ID", "CN", "HK", "MO" };


        private void Radio_Checked(object sender, RoutedEventArgs e)
        {
            var radio = sender as RadioButton;
            messageBird = Convert.ToBoolean(radio.Tag);
            //for debugging
            //if (textBlock != null)
            //{
            //    textBlock.Text = messageBird.ToString();
            //}          
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Excel files (*.csv;*.xlsx;*.xls)|*.csv;*.xlsx;*.xls";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (openFileDialog.ShowDialog() == true)
                FileName.Content = openFileDialog.SafeFileName;

            if (uploadedFilePath != null && uploadedFilePath != openFileDialog.FileName)
                ReleaseMemory(range, MySheet, MyBook, MyApp);

            uploadedFilePath = openFileDialog.FileName;

            // just to make sure
            MyBook = null;
            MyApp = null;
            MySheet = null;

            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(uploadedFilePath);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            Locale = MySheet.Name.ToUpperInvariant();
            lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            range = MySheet.UsedRange;
            rowsLabel.Content = lastRow.ToString() + " rows";

            if (!markets.Any(Locale.ToUpperInvariant().Contains))
            {
                marketLabel.Foreground = new SolidColorBrush(Colors.Red);
                marketLabel.Content = "Remember to rename the first sheet with your two-letter market code";
                startBtn.IsEnabled = false;
            }
            else
            {
                marketLabel.Foreground = new SolidColorBrush(Colors.Black);
                marketLabel.Content = "Market: " + Locale;
                startBtn.IsEnabled = true;
            }
        }

        private void start_Click(object sender, RoutedEventArgs e)
        {
            progressBar.Value = 0;
            startBtn.IsEnabled = false;

            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += ParseExcelFile;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.RunWorkerAsync();
        }

        public void ParseExcelFile(object sender, DoWorkEventArgs e)
        {
            try
            {
                //export the Excel into a DataTable for easier manipulation
                dt = ExceltoDataTable(sender);
                newFileName = uploadedFilePath + ".csv";

                //Process Datatable
                dt = ProcessData(sender);

                //Export the Datatable into a CSV file
                CreateCSV(sender);
            }
            catch (Exception ex)
            {
                //log it
                textBlock.Text = ex.InnerException.ToString();
            }
            finally
            {
                //clean up
                ReleaseMemory(range, MySheet, MyBook, MyApp);
            }

        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            if (e.UserState != null && e.UserState.GetType() == typeof(string)) ProgressBarLabel.Content = e.UserState;
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressBarLabel.Content = "Done!";
        }

        private DataTable ExceltoDataTable(object sender)
        {
            DataTable DT = new DataTable();
            object[,] data = range.Value2;
            int index = 0;

            //update UI
            (sender as BackgroundWorker).ReportProgress(0, "Loading data...");

            // Create new Column in DataTable
            try
            {
                for (int cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {
                    int progressPercentage = Convert.ToInt32(((double)cCnt / (range.Columns.Count - 1)) * 100);
                    (sender as BackgroundWorker).ReportProgress(progressPercentage);

                    if ((data[1, cCnt] == null) || (data[1, cCnt].ToString() == string.Empty))
                    {
                        continue;
                    }

                    var columnName = (data[1, cCnt]).ToString().Trim().ToLowerInvariant();

                    if (!arr.Any(columnName.Contains))
                    {
                        continue;
                    }
                    else
                    {
                        DataColumn Column = new DataColumn();
                        Column.DataType = Type.GetType("System.String");
                        if (columnName == SubscriberKey ||
                            columnName == customerID ||
                            columnName == customer_ID)
                        {
                            Column.ColumnName = SubscriberKey;
                            columnName = SubscriberKey;
                        }
                        else if (columnName == MobilePhone || columnName == Mobile_Phone || columnName == MobileNumber)
                        {
                            Column.ColumnName = MobilePhone;
                            columnName = MobilePhone;
                        }
                        else
                        {
                            Column.ColumnName = columnName;
                        }

                        DT.Columns.Add(Column);
                        index++;
                    }

                    // Create row for Data Table
                    for (int rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                    {

                        string CellVal = String.Empty;
                        try
                        {
                            CellVal = data[rCnt, cCnt] == null ? String.Empty : data[rCnt, cCnt].ToString();
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }

                        DataRow Row;

                        // Add to the DataTable
                        if (index == 1)
                        {
                            Row = DT.NewRow();
                            Row[columnName] = CellVal;
                            DT.Rows.Add(Row);
                        }
                        else
                        {
                            Row = DT.Rows[rCnt - 1];
                            Row[columnName] = CellVal;
                        }
                    }
                }

                //create a column for the Locale
                var col = new DataColumn();
                col.DataType = Type.GetType("System.String");
                col.ColumnName = "Locale";
                DT.Columns.Add(col);
                DT.Rows[0][col] = "Locale";

                return DT;
            }
            catch (Exception ex)
            {
                textBlock.Text = ex.InnerException.ToString();
            }
            return null;
        }

        private DataTable ProcessData(object sender)
        {
            try
            {
                for (int i = dt.Rows.Count - 1; i >= 0; i--)
                {
                    int progressPercentage = Convert.ToInt32(100 - (((double)i / (dt.Rows.Count - 1)) * 100));
                    (sender as BackgroundWorker).ReportProgress(progressPercentage, "Cleaning...");

                    DataRow row = dt.Rows[i];

                    //don't process first row since they are only the headers
                    if (row == dt.Rows[0]) continue;

                    //add the locale column
                    row["Locale"] = Locale;

                    switch (Locale.ToUpperInvariant())
                    {
                        case "JP":
                            // row[SubscriberKey] = "JPT-CUS-" + row[SubscriberKey];
                            row[MobilePhone] = ProcessNumber(row[MobilePhone].ToString(), "81");
                            break;
                        case "TH":
                            row[MobilePhone] = ProcessNumber(row[MobilePhone].ToString(), "66");
                            break;
                        case "ID":
                            row[MobilePhone] = ProcessNumber(row[MobilePhone].ToString(), "62");
                            break;
                        case "TW":
                            row[MobilePhone] = ProcessNumber(row[MobilePhone].ToString(), "886");
                            break;
                        case "KR":
                            row[MobilePhone] = ProcessNumber(row[MobilePhone].ToString(), "82");
                            break;
                        case "VN":
                            row[MobilePhone] = ProcessNumber(row[MobilePhone].ToString(), "84");
                            break;
                        case "CN":
                            row[MobilePhone] = ProcessNumber(row[MobilePhone].ToString(), "86");
                            break;
                        case "HK":
                            row[MobilePhone] = ProcessNumber(row[MobilePhone].ToString(), "852");
                            break;
                        case "MO":
                            row[MobilePhone] = ProcessNumber(row[MobilePhone].ToString(), "853");
                            break;
                        default:
                            throw new Exception("Locale not found on Sheet Name");
                    }
                    if (row[MobilePhone].ToString() == String.Empty) dt.Rows.Remove(row);
                }

                return dt;
            }
            catch (Exception ex)
            {
                textBlock.Text = ex.Message;
            }
            return null;
        }

        public void CreateCSV(object sender)
        {
            if (!Directory.Exists(Path.GetDirectoryName(newFileName)))
                throw new DirectoryNotFoundException($"Destination folder not found: {newFileName}");

            //string delimiter = ",";
            (sender as BackgroundWorker).ReportProgress(0, "Saving...");

            StringBuilder sb = new StringBuilder();
            Regex regex = new Regex(@"^[\,]+$");

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow row = dt.Rows[i];
                int progressPercentage = Convert.ToInt32(((double)i / (dt.Rows.Count - 1)) * 100);
                (sender as BackgroundWorker).ReportProgress(progressPercentage);

                StringBuilder sbtemp = new StringBuilder();
                for (int h = 0; h < dt.Columns.Count; h++)
                {
                    sbtemp.Append(row[h].ToString() + ",");
                }
                if (!regex.IsMatch(sbtemp.ToString()))
                {
                    sb.Append(sbtemp.ToString() + ",");
                    sb.AppendLine();
                }
            }

            File.WriteAllText(newFileName, sb.ToString(), Encoding.UTF8);
        }

        private void ReleaseMemory(Excel.Range xlRange, Excel.Worksheet xlWorksheet, Excel.Workbook xlWorkbook, Excel.Application xlApp)
        {
            //Garbage Collector
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        private string ProcessNumber(string number, string prepend)
        {
            string cleanNumber;
            string plusPrepend = '+' + prepend;
            if (number.StartsWith(prepend) || number.StartsWith(plusPrepend)) { cleanNumber = string.Empty; }
            else cleanNumber = prepend;

            var regex = new Regex(@"[0-9]+", RegexOptions.CultureInvariant);

            if (number != string.Empty && regex.IsMatch(number))
            {
                MatchCollection collection = regex.Matches(number);

                foreach (Match match in collection)
                {
                    if (match.Index == 0 && match.ToString().StartsWith("0"))
                    {
                        cleanNumber += match.ToString().Substring(1);
                    }
                    else cleanNumber += match.ToString();
                }
                if (messageBird)
                {
                    //the + won't show in excel but it is there in the csv file
                    cleanNumber = "+" + cleanNumber;
                }
            }
            else return string.Empty;

            return cleanNumber;
        }

        private Version getRunningVersion()
        {
            try
            {
                return ApplicationDeployment.IsNetworkDeployed ? ApplicationDeployment.CurrentDeployment.CurrentVersion
                : Assembly.GetExecutingAssembly().GetName().Version;
            }
            catch (Exception)
            {
                return Assembly.GetExecutingAssembly().GetName().Version;
            }
        }
    }
}
