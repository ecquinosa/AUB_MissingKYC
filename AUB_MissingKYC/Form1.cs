using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;

namespace AUB_MissingKYC
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string configFile = AppDomain.CurrentDomain.BaseDirectory + "config";
        private DataTable dt = null;
        private DataTable dtDateGroup = null;
        private NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private DAL dal = new DAL();
        private Config config;
        private string missingFilesFolder = "";

        private string[] refNoColumns = { "[Reference No.]", "refno" };
        private string[] midColumns = { "MID", "idno" };
        private string refNoColumn = "";
        private string midColumn = "";

        private void Form1_Load(object sender, EventArgs e)
        {
            //Init();
            //SyncDirectories();
            //return;    

            //PopulateExceptionList();
            //SFTP sftp = new SFTP(config);
            //sftp.GetListOfFile("", lstException);
            //sftp = null;
            //return;
            
            logger.Info("Application started");
            if (!Init())
            {
                MessageBox.Show("Failed to initialize config. Check error log for details.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnGenerate.Enabled = false;
            }
        }

        public void PrepExcelData(string FilePath)
        {
            try
            {
                if (dtDateGroup == null)
                {
                    dtDateGroup = new DataTable();
                    dtDateGroup.Columns.Add("Date", Type.GetType("System.String"));
                    dtDateGroup.Columns.Add("Count", Type.GetType("System.String"));
                    dtDateGroup.Columns.Add("IsSelected", Type.GetType("System.Boolean"));
                }
                else dtDateGroup.Clear();


                FileInfo existingFile = new FileInfo(FilePath);
                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    // get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    dt = ToDataTable(package.Workbook.Worksheets[1], true);

                    foreach (DataRow r in dt.Rows)
                    {
                        if (r["idno"].ToString().Trim() != "")
                        {
                            string d = r["date"].ToString();
                            r["date"] = d.Substring(0, 4) + "-" + d.Substring(4, 2) + "-" + d.Substring(6, 2);
                            r.AcceptChanges();
                        }
                        //Console.WriteLine(r["idno"].ToString());
                        //if (r["idno"].ToString().Trim() == "121085790074")
                        //{
                        //    Console.WriteLine(r["idno"].ToString());
                        //}                        
                    }

                    var groupedData = from b in dt.AsEnumerable()
                                      group b by b.Field<string>("Date") into g
                                      select new
                                      {
                                          Date = g.Key,
                                          Count = g.Count()
                                      };

                    foreach (var gd in groupedData)
                    {
                        if (Convert.ToDateTime(gd.Date) > config.ReportStartDate)
                        {
                            DataRow rw = dtDateGroup.NewRow();
                            rw[0] = gd.Date;
                            rw[1] = gd.Count;
                            rw[2] = false;
                            dtDateGroup.Rows.Add(rw);
                        }
                    }

                }

                grid.DataSource = dtDateGroup;
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw new Exception();
            }
        }

        public DataTable ToDataTable(ExcelWorksheet ws, bool hasHeaderRow = true)
        {
            var tbl = new DataTable();
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                tbl.Columns.Add(hasHeaderRow ?
                    firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            var startRow = hasHeaderRow ? 2 : 1;
            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                var row = tbl.NewRow();
                foreach (var cell in wsRow) row[cell.Start.Column - 1] = cell.Text;
                tbl.Rows.Add(row);
            }
            return tbl;
        }

        private void btnExcelFile_Click(object sender, EventArgs e)
        {            
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtExcelFile.Text = ofd.FileName;
                btnGenerate.Enabled = false;

                this.Cursor = Cursors.WaitCursor;
                PrepExcelData(txtExcelFile.Text);
                this.Cursor = Cursors.Default;

                btnGenerate.Enabled = true;
            }
            ofd.Dispose();
            ofd = null;
        }

        List<String> dateList = new List<String>();
        private void grid_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string selectedDate = grid.Rows[e.RowIndex].Cells[0].Value.ToString();
            //if (dateList.Find(d => d == selectedDate).Count() > 0) dateList.Remove(selectedDate);


            bool isExist = Convert.ToBoolean(dtDateGroup.Select("Date='" + selectedDate + "'")[0]["IsSelected"]);
            if (isExist) dateList.Remove(selectedDate); else dateList.Add(selectedDate);
            dtDateGroup.Select("Date='" + selectedDate + "'")[0]["IsSelected"] = !isExist;

            BindGrids();
        }

        private DataTable FilteredTable()
        {
            System.Text.StringBuilder sb = new StringBuilder();

            foreach (var d in dateList)
            {
                if (sb.ToString() == "") sb.Append(d); else sb.Append("','" + d);
            }

            DataTable temp = null;
            if (dt.Select("Date IN ('" + sb.ToString() + "')").Count() > 0)
                temp = dt.Select("Date IN ('" + sb.ToString() + "')").CopyToDataTable();
            else
                temp = dt.Clone();

            return temp;
        }

        private void BindGrids()
        {
            DataTable temp = FilteredTable();
            grid2.DataSource = temp;
            lblTotal.Text = "TOTAL: " + temp.DefaultView.Count.ToString("N0");
            grid.DataSource = dtDateGroup;
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            this.Cursor = Cursors.WaitCursor;

            btnMove.Enabled = false;
            btnTransferToSftp.Enabled = false;

            if (!dt.Columns.Contains("Remark")) dt.Columns.Add("Remark", Type.GetType("System.String"));
            if (!dt.Columns.Contains("Zipped")) dt.Columns.Add("Zipped", Type.GetType("System.String"));
            if (!dt.Columns.Contains("Sftp")) dt.Columns.Add("Sftp", Type.GetType("System.String"));
            bank_ws.ACC_MS_WEBSERVICE ws = null;

            try
            {
                if (chkForce.Checked) ws = new bank_ws.ACC_MS_WEBSERVICE();
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
            }

            foreach (var colName in refNoColumns) if (isColumnNameExist(colName, dt)) refNoColumn = colName;
            foreach (var colName in midColumns) if (isColumnNameExist(colName, dt)) midColumn = colName;

            if (refNoColumn == "")
            {
                MessageBox.Show("Unable to find columns with headers " + String.Join(",", refNoColumns), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (midColumn == "")
            {
                MessageBox.Show("Unable to find columns with headers " + String.Join(",", midColumns), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //missingFilesFolder = string.Format(@"{0}\{1}_TESTONLY", config.BankForTransferFolder, DateTime.Now.ToString("MMddyyyyhhmmss"));
            missingFilesFolder = string.Format(@"{0}\{1}_MISSINGFILES", config.BankForTransferFolder, DateTime.Now.ToString("MMddyyyy"));
            if (!Directory.Exists(missingFilesFolder)) Directory.CreateDirectory(missingFilesFolder);

            try
            {
                foreach (DataRow rw in FilteredTable().Rows)
                {                   

                    //string missingFilesFolder = string.Format(@"{0}\{1}_MISSINGFILES", config.BankForTransferFolder, DateTime.Now.ToString("MMddyyyy"));
                  
                    if (dal.GetEntryDateByMID(rw[midColumn].ToString().Trim()))
                    {
                        //if (dal.ObjectResult == null) rw["Remark"] = "No data in member table";
                        if (dal.TableResult.DefaultView.Count == 0) rw["Remark"] = "No data in member table";
                        else
                        {
                            DateTime entryDate = Convert.ToDateTime(dal.TableResult.Rows[0]["EntryDate"]); //Convert.ToDateTime(dal.ObjectResult.ToString());
                            DateTime startDate = entryDate.AddDays(-4);
                            DateTime endDate = entryDate.AddDays(4);
                            DateTime condDate = startDate;

                            string zipFolder = "";
                            string zipFolder2 = "";

                            while (condDate <= endDate)
                            {
                                zipFolder = string.Format(@"{0}\{1}\{2}.zip", config.BankDoneFolder, condDate.ToString("yyyy-MM-dd"), rw[midColumn].ToString().Trim());
                                zipFolder2 = zipFolder;
                                if (File.Exists(zipFolder))
                                {
                                    zipFolder = "Transferred to " + condDate.ToShortDateString() + " and " + missingFilesFolder.Substring(missingFilesFolder.LastIndexOf("\\") + 1) + " sftp folders";

                                    break;
                                }
                                condDate = condDate.AddDays(1);
                            }

                            if (zipFolder != "")
                                rw["Remark"] = zipFolder;
                            else rw["Remark"] = entryDate.ToShortDateString();

                            if (rw["Remark"].ToString().Contains("\\DONE\\")) rw["Remark"] = "Transferred to " + missingFilesFolder.Substring(missingFilesFolder.LastIndexOf("\\") + 1) + " sftp folder";


                            if (File.Exists(zipFolder2))
                            {
                                //if (!Directory.Exists(missingFilesFolder)) Directory.CreateDirectory(missingFilesFolder);
                                File.Copy(zipFolder2, string.Format(@"{0}\{1}", missingFilesFolder, Path.GetFileName(zipFolder2)));
                            }
                            else
                            {
                                try
                                {
                                    if (ws != null)
                                    {
                                        var response = ws.ManualPackUpData(dal.TableResult.Rows[0]["RefNum"].ToString(), "");
                                        if (!response.IsSuccess)
                                            logger.Error("Failed to extract RefNum " + dal.TableResult.Rows[0]["RefNum"].ToString() + " MID " + rw[midColumn].ToString().Trim() + ". ManualPackUpData error " + response.ErrorMessage);
                                    }
                                    }
                                catch (Exception ex)
                                {
                                    logger.Error("Failed to extract RefNum " + dal.TableResult.Rows[0]["RefNum"].ToString() + " MID " + rw[midColumn].ToString().Trim() + ". ManualPackUpData error " + ex.Message);
                                }
                            }
                        }
                    }
                    else
                    {
                        rw["Remark"] = "Db error";
                        logger.Error("MID " + rw[midColumn].ToString().Trim() + ". Error " + dal.ErrorMessage);
                    }

                    rw.AcceptChanges();                    

                    //dt.Select("[Reference No.]='" + rw["Reference No."].ToString().Trim() + "' AND MID='" + rw["MID"].ToString().Trim() + "'")[0]["Remark"] = rw["Remark"].ToString().Trim();
                    //dt.Select(refNoColumn +  "='" + rw[refNoColumn].ToString().Trim() + "' AND MID='" + rw[midColumn].ToString().Trim() + "'")[0]["Remark"] = rw["Remark"].ToString().Trim();
                    dt.Select(string.Format("{0}='{1}' AND {2}='{3}'",refNoColumn, rw[refNoColumn].ToString().Trim(),midColumn, rw[midColumn].ToString().Trim()))[0]["Remark"] = rw["Remark"].ToString().Trim();
                }

                ZipAndMoveFiles();
                SyncDirectories();
            }
            catch (Exception ex)
            {
                logger.Error("For loop error. " + ex.Message);
            }

            if (ws != null)
            {
                ws.Dispose();
                ws = null;
            }

            BindGrids();

            btnReport.PerformClick();

            MessageBox.Show("Process is complete", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            btnMove.Enabled = true;
            btnTransferToSftp.Enabled = true;

            this.Enabled = true;
            this.Cursor = Cursors.Default;
        }

        private void btnGenerate_Clickv2(object sender, EventArgs e)
        {
            this.Enabled = false;
            this.Cursor = Cursors.WaitCursor;

            if (!dt.Columns.Contains("Remark")) dt.Columns.Add("Remark", Type.GetType("System.String"));
            bank_ws.ACC_MS_WEBSERVICE ws = null;

            try
            {
                if (chkForce.Checked) ws = new bank_ws.ACC_MS_WEBSERVICE();
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
            }

            try
            {
                foreach (DataRow rw in FilteredTable().Rows)
                {
                    string[] refNoColumns = { "[Reference No.]", "refno" };
                    string[] midColumns = { "MID", "idno" };
                    string refNoColumn = "";
                    string midColumn = "";

                    foreach (var colName in refNoColumns) if (isColumnNameExist(colName, dt)) refNoColumn = colName;
                    foreach (var colName in refNoColumns) if (isColumnNameExist(colName, dt)) midColumn = colName;

                    if (dal.GetEntryDateByMID(rw["MID"].ToString()))
                    {
                        //if (dal.ObjectResult == null) rw["Remark"] = "No data in member table";
                        if (dal.TableResult.DefaultView.Count == 0) rw["Remark"] = "No data in member table";
                        else
                        {
                            DateTime entryDate = Convert.ToDateTime(dal.TableResult.Rows[0]["EntryDate"]); //Convert.ToDateTime(dal.ObjectResult.ToString());
                            DateTime startDate = entryDate.AddDays(-4);
                            DateTime endDate = entryDate.AddDays(4);
                            DateTime condDate = startDate;

                            string zipFolder = "";
                            string zipFolder2 = "";

                            while (condDate <= endDate)
                            {
                                zipFolder = string.Format(@"{0}\{1}\{2}.zip", config.BankDoneFolder, condDate.ToString("yyyy-MM-dd"), rw["MID"].ToString());
                                zipFolder2 = zipFolder;
                                if (File.Exists(zipFolder))
                                {
                                    zipFolder = "Transferred on " + condDate.ToShortDateString();

                                    break;
                                }
                                condDate = condDate.AddDays(1);
                            }

                            if (zipFolder != "")
                                rw["Remark"] = zipFolder;
                            else rw["Remark"] = entryDate.ToShortDateString();

                            string missingFilesFolder = string.Format(@"D:\ACCPAGIBIGPH3\AUB\PACKUPDATA\FOR_TRANSFER\{0}_MISSINGFILES", DateTime.Now.ToString("MMddyyyy"));

                            if (File.Exists(zipFolder2))
                            {
                                if (!Directory.Exists(missingFilesFolder)) Directory.CreateDirectory(missingFilesFolder);
                                File.Copy(zipFolder2, string.Format(@"{0}\{1}", missingFilesFolder, Path.GetFileName(zipFolder2)));
                            }
                            else
                            {
                                try
                                {
                                    var response = ws.ManualPackUpData(dal.TableResult.Rows[0]["RefNum"].ToString(), "");
                                    if (!response.IsSuccess)
                                        logger.Error("Failed to extract RefNum " + dal.TableResult.Rows[0]["RefNum"].ToString() + " MID " + rw["MID"].ToString() + ". ManualPackUpData error " + response.ErrorMessage);
                                }
                                catch (Exception ex)
                                {
                                    logger.Error("Failed to extract RefNum " + dal.TableResult.Rows[0]["RefNum"].ToString() + " MID " + rw["MID"].ToString() + ". ManualPackUpData error " + ex.Message);
                                }
                            }
                        }
                    }
                    else
                    {
                        rw["Remark"] = "Db error";
                        logger.Error("MID " + rw["MID"].ToString() + ". Error " + dal.ErrorMessage);
                    }

                    rw.AcceptChanges();

                    //dt.Select("[Reference No.]='" + rw["Reference No."].ToString().Trim() + "' AND MID='" + rw["MID"].ToString().Trim() + "'")[0]["Remark"] = rw["Remark"].ToString().Trim();
                    dt.Select(refNoColumn + "='" + rw["Reference No."].ToString().Trim() + "' AND MID='" + rw["MID"].ToString().Trim() + "'")[0]["Remark"] = rw["Remark"].ToString().Trim();
                }
            }
            catch (Exception ex)
            {
                logger.Error("For loop error. " + ex.Message);
            }

            if (ws != null)
            {              
                ws.Dispose();
                ws = null;
            }

            BindGrids();

            btnReport.PerformClick();

            MessageBox.Show("Process is complete", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.Enabled = true;
            this.Cursor = Cursors.Default;
        }

        private bool isColumnNameExist(string columnName, DataTable table)
        {
            DataColumnCollection columns = table.Columns;
            if (columns.Contains(columnName)) return true; else return false;          
        }

        private void GenerateReport()
        {
            this.Enabled = false;
            this.Cursor = Cursors.WaitCursor;

            DataTable temp = FilteredTable();

            var groupedData = from b in temp.AsEnumerable()
                              group b by new
                              {
                                  dateValue = b.Field<string>("Date"),
                                  remarkVal = b.Field<string>("Remark")
                              } into gcs
                              select new
                              {
                                  Date = gcs.Key.dateValue,
                                  Remark = gcs.Key.remarkVal,
                                  Count = gcs.Count(),

                              };

            StringBuilder sbSummary = new StringBuilder();
            StringBuilder sbDetails = new StringBuilder();
            foreach (var gd in groupedData) sbSummary.AppendLine(gd.Date + "\t" + gd.Count.ToString() + "\t" + gd.Remark);

            foreach (DataRow rw in temp.Rows)
            {
                sbDetails.AppendLine(rw[0].ToString() + "\t" + rw[1].ToString() + "\t" + rw[2].ToString() + "\t" + rw[3].ToString());
            }


            string fileTimeStamp = DateTime.Now.ToString("hhmmss");

            File.WriteAllText(string.Format(@"{0}\logs\ListSum{1}.txt", AppDomain.CurrentDomain.BaseDirectory, fileTimeStamp), sbSummary.ToString());
            File.WriteAllText(string.Format(@"{0}\logs\ListDtl{1}.txt", AppDomain.CurrentDomain.BaseDirectory, fileTimeStamp), sbDetails.ToString());

            MessageBox.Show("Process is complete", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.Enabled = true;
            this.Cursor = Cursors.Default;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            NLog.LogManager.Shutdown();
        }

        private bool Init()
        {
            try
            {
                //check if file exists
                if (!File.Exists(configFile))
                {
                    logger.Error("Config file is missing");
                    return false;
                }

                try
                {
                    config = new Config();
                    var configData = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Config>>(File.ReadAllText(configFile));
                    config = configData[0];
                    dal.ConStr = config.DbaseConStr;
                }
                catch (Exception ex)
                {
                    logger.Error("Error reading config file. Runtime catched error " + ex.Message);
                    return false;
                }

                //check dbase connection
                if (!dal.IsConnectionOK())
                {
                    logger.Error("Connection to database failed. " + dal.ErrorMessage);
                    return false;
                }

            }
            catch (Exception ex)
            {
                logger.Error("Runtime catched error " + ex.Message);
                return false;
            }

            return true;
        }

        private void btnReport_Click(object sender, EventArgs e)
        {
            GenerateReport();
        }

        private void ZipAndMoveFiles()
        {
            foreach (DataRow rw in FilteredTable().Rows)
            {
                string folder = string.Format(@"{0}\{1}", config.KycDumpfileFolder, rw[midColumn].ToString().Trim());
                if (Directory.Exists(folder))
                {
                    string zipFile = "";
                    if (FileCompression.Compress(folder, folder, ref zipFile))
                    {                        
                        dt.Select(string.Format("{0}='{1}' AND {2}='{3}'", refNoColumn, rw[refNoColumn].ToString().Trim(), midColumn, rw[midColumn].ToString().Trim()))[0]["Zipped"] = DateTime.Now.ToString();
                        File.Move(zipFile, string.Format(@"{0}\{1}",missingFilesFolder,Path.GetFileName(zipFile)));
                        Directory.Delete(folder, true);
                    }
                }
            }           
        }

        private void SyncDirectories()
        {
            SFTP sftp = new SFTP(config);
            string errMsg = "";
            int totalTransferred = 0;
            sftp.SynchronizeDirectories(ref errMsg, ref totalTransferred);
        }

        private void TempProcess()
        {
            //DAL localDAL = new DAL();

            using (StreamReader sr = new StreamReader("sftpFiles.txt"))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    if (line.Trim() != "")
                    {
                        //if (!dal.ExecuteQuery(line)) MessageBox.Show(dal.ErrorMessage);
                        dal.ExecuteQuery(line);
                    }
                }
            }

            //localDAL.Dispose();
            //localDAL = null;
        }

        private void btnMove_Click(object sender, EventArgs e)
        {
            ZipAndMoveFiles();
        }

        private void btnTransferToSftp_Click(object sender, EventArgs e)
        {
            SyncDirectories();
        }


        private List<string> lstException = new List<String>();

        private void PopulateExceptionList()
        {
            using (StreamReader sr = new StreamReader("exception.txt"))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    if (line.Trim() != "")
                    {
                        lstException.Add(line.Trim());
                    }
                }
                sr.Dispose();
                sr.Close();
            }
        }
    }
}

