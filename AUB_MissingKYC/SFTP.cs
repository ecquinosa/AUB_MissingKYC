using System;
using System.Collections.Generic;
using WinSCP;
using System.IO;

namespace AUB_MissingKYC
{
    class SFTP
    {

        public string ErrorMessage { get; set; }

        private delegate void dlgtProcess();
        private Config config;
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        public SFTP(Config config)
        {
            this.config = config;
        }

        private SessionOptions sessionOptions()
        {
            return new SessionOptions
            {
                Protocol = Protocol.Sftp,
                HostName = config.SftpHost,
                UserName = config.SftpUser,
                Password = config.SftpPass,
                PortNumber = Convert.ToInt32(config.SftpPort),
                SshHostKeyFingerprint = config.SftpSshHostKeyFingerprint
            };
        }

        //public bool DownloadFiles()
        //{
        //    try
        //    {  
        //        using (Session session = new Session())
        //        {
        //            // Connect
        //            session.Open(sessionOptions());

        //            // Download files
        //            TransferOptions transferOptions = new TransferOptions();
        //            transferOptions.TransferMode = TransferMode.Binary;

        //            TransferOperationResult transferResult;
        //            transferResult =
        //                session.GetFiles(config.SftpRemotePath, string.Format(@"{0}\",config.DownloadFolder), false, transferOptions);

        //            // Throw on any error
        //            transferResult.Check();

        //            // Print results
        //            foreach (TransferEventArgs transfer in transferResult.Transfers)
        //            {
        //                Console.WriteLine(DateTime.Now.ToString("MM/dd/yy hh:mm:ss ") + "Download of {0} succeeded", transfer.FileName);
        //            }
        //        }                

        //        return true;
        //    }
        //    catch (Exception e)
        //    {
        //        ErrorMessage = e.Message;
        //        return false;
        //    }
        //}

        //public bool Upload_SFTP_Files(Cstring path, bool IsZip, ref string errMsg)
        //{
        //    try
        //    {
        //        int intFileCount = Directory.GetFiles(SFTP_LOCALPATH).Length;

        //        if (intFileCount == 0)
        //        {
        //            errMsg = string.Format("[Upload] {0} is empty. No file to push.", SFTP_LOCALPATH);                  
        //            return false;
        //        }             

        //        using (Session session = new Session())
        //        {                                 
        //            session.DisableVersionCheck = true;

        //            session.Open(sessionOptions());

        //            // Upload files
        //            TransferOptions transferOptions = new TransferOptions();
        //            transferOptions.TransferMode = TransferMode.Binary;
        //            //transferOptions.ResumeSupport.State = TransferResumeSupportState.Smart;                  

        //            //transferOptions.PreserveTimestamp = false;

        //            //Console.Write(AppDomain.CurrentDomain.BaseDirectory);
        //            string remotePath = SFTP_REMOTEPATH_ZIP;
        //            if (!IsZip) remotePath = SFTP_REMOTEPATH_PAGIBIGMEMU;

        //             TransferOperationResult transferResult = null;
        //            if (File.Exists(path))
        //            {
        //                {
        //                    if (!session.FileExists(remotePath + Path.GetFileName(path)))
        //                    {
        //                        transferResult = session.PutFiles(string.Format(@"{0}*", path), remotePath, false, transferOptions);
        //                    }

        //                    else
        //                    {
        //                        errMsg = string.Format("Upload_SFTP_Files(): Remote file exist " + Path.GetFileName(path));                               
        //                        return false;
        //                    }
        //                }
        //            }
        //              else

        //                transferResult = session.PutFiles(string.Format(@"{0}\*", SFTP_LOCALPATH), remotePath, false, transferOptions);


        //                // Throw on any error
        //                transferResult.Check();

        //                // Print results
        //                foreach (TransferEventArgs transfer in transferResult.Transfers)
        //                {
        //                    //Console.WriteLine(TimeStamp() + Path.GetFileName(transfer.FileName) + " transferred successfully");
        //                    //string strFilename = Path.GetFileName(transfer.FileName);
        //                    //File.Delete(transfer.FileName);
        //                }                        
        //            }

        //        //Console.WriteLine("Success sftp transfer " + path);
        //        //System.Threading.Thread.Sleep(100);

        //        return true;

        //    }                            
        //    catch (Exception ex)
        //    {
        //        errMsg = string.Format("Upload_SFTP_Files(): Runtime error {0}", ex.Message);
        //        Console.WriteLine(errMsg);
        //        //Utilities.WriteToRTB(errMsg, ref rtb, ref tssl);
        //        return false;
        //    }
        //}

        //private static string BANK_REPO = "";
        //private static System.Text.StringBuilder sbDone = new System.Text.StringBuilder();
        private static int TotalSftpTransfer;

        public bool SynchronizeDirectories(ref string errMsg, ref int _TotalSftpTransfer)
        {
            try
            {
                 string forTransferFolder = config.BankForTransferFolder;

                int intFileCount = Directory.GetFiles(forTransferFolder).Length;

                //if (intFileCount == 0)
                //{
                //    errMsg = string.Format("[Upload] {0} is empty. No file to push.", forTransferFolder);                    
                //    return false;
                //}   

                dal.ConStr = config.DbaseConStr;

                using (Session session = new Session())
                {
                    session.DisableVersionCheck = true;

                    TransferOptions transferOptions = new TransferOptions();
                    transferOptions.TransferMode = TransferMode.Binary;
                    transferOptions.FilePermissions = null;
                    transferOptions.PreserveTimestamp = false;


                    // Will continuously report progress of synchronization
                    session.FileTransferred += FileTransferred;

                    // Connect
                    session.Open(sessionOptions());                    

                    // Synchronize files
                    SynchronizationResult synchronizationResult;
                    synchronizationResult = session.SynchronizeDirectories(SynchronizationMode.Remote, forTransferFolder, config.SftpRemotePath, false, false, SynchronizationCriteria.None, transferOptions);

                    // Throw on any error
                    synchronizationResult.Check();
                }

                _TotalSftpTransfer = TotalSftpTransfer;

                return true;
            }
            catch (Exception ex)
            {
                errMsg = string.Format("SynchronizeDirectories(): Runtime error {0}", ex.Message);
                logger.Error(errMsg);
                Utilities.ShowWarningMessage(ex.Message, "SFTP");
                Console.WriteLine(errMsg);
                return false;
            }
        }

        private DAL dal = new DAL();

        private void FileTransferred(object sender, TransferEventArgs e)
        {
            string msg = "";
            if (e.Error == null)
            {
                msg = string.Format("{0}Upload of {1} succeeded", Utilities.TimeStamp(), Path.GetFileName(e.FileName));
                Console.WriteLine(msg);
                logger.Info(msg);

                string[] arr = e.FileName.Split('\\');
                string sourceFolder = "";
                try
                {
                    sourceFolder = arr[arr.Length - 2];
                }
                catch { }

                string destiFile = e.FileName.Replace("FOR_TRANSFER", "DONE");
                if (File.Exists(destiFile))
                {
                    string destiFileExisting = string.Format("{0}_{1}.zip", Path.GetFileNameWithoutExtension(destiFile), new FileInfo(destiFile).CreationTime.ToString("yyyyMMdd_hhmmss"));
                    if (!File.Exists(destiFileExisting)) File.Move(destiFile, destiFileExisting);
                }

                try
                {
                    if (!Directory.Exists(config.BankDoneFolder + "\\" + sourceFolder)) Directory.CreateDirectory(config.BankDoneFolder + "\\" + sourceFolder);
                    File.Move(e.FileName, destiFile);
                }
                catch { }  
                
                dal.InsertSFTP(Path.GetFileName(e.FileName).Replace(".zip", ""), sourceFolder, DateTime.Now);

                TotalSftpTransfer += 1;
                System.Threading.Thread.Sleep(100);
            }
            else
            {
                msg = string.Format("{0}Upload of {1} failed: {2}", Utilities.TimeStamp(), Path.GetFileName(e.FileName), e.Error);
                Console.WriteLine(msg);
                logger.Error(msg);
            }

            if (e.Chmod != null)
            {
                if (e.Chmod.Error == null)
                {
                    msg = string.Format("{0}Permissions of {1} set to {2}", Utilities.TimeStamp(), Path.GetFileName(e.Chmod.FileName), e.Chmod.FilePermissions);
                    Console.WriteLine(msg);
                    logger.Info(msg);
                }
                else
                {
                    msg = string.Format("{0}Setting permissions of {1} failed: {2}", Utilities.TimeStamp(), Path.GetFileName(e.Chmod.FileName), e.Chmod.Error);
                    Console.WriteLine(msg);
                    logger.Error(msg);
                }
            }
            else
            {
                //Console.WriteLine("{0}Permissions of {1} kept with their defaults", TimeStamp(), e.Destination);
            }

            if (e.Touch != null)
            {
                if (e.Touch.Error == null)
                {
                    msg = string.Format("{0}Timestamp of {1} set to {2}", Utilities.TimeStamp(), Path.GetFileName(e.Touch.FileName), e.Touch.LastWriteTime);
                    Console.WriteLine(msg);
                    logger.Error(msg);                   
                    
                }
                else
                {
                    msg = string.Format("{0}Setting timestamp of {1} failed: {2}", Utilities.TimeStamp(), Path.GetFileName(e.Touch.FileName), e.Touch.Error);
                    Console.WriteLine(msg);
                    logger.Error(msg);
                }
            }
            else
            {
                // This should never happen during "local to remote" synchronization                
                msg = string.Format("{0}Timestamp of {1} kept with its default (current time)", Utilities.TimeStamp(), e.Destination);
                Console.WriteLine(msg);
                logger.Error(msg);
            }
        }

        public bool GetListOfFile(string dir, List<string> lstException)
        {
            try
            {            
                using (Session session = new Session())
                {
                    session.DisableVersionCheck = true;
                    // Connect
                    session.Open(sessionOptions());                
                 
                    string parentDir = config.SftpRemotePath;

                    RemoteDirectoryInfo directory =
                        session.ListDirectory(parentDir);

                    SaveFTPFile(string.Format("{0}|Folder", parentDir));

                    foreach (RemoteFileInfo fileInfo in directory.Files)
                    {
                        switch (fileInfo.Name)
                        {
                            case ".":
                            case "..":
                                break;
                            default:
                                if (fileInfo.IsDirectory)
                                {
                                    string curDir = parentDir + fileInfo.Name;

                                    var match = lstException.Find(stringToCheck => stringToCheck.Contains(curDir));

                                    if (match == null)
                                    {
                                        SaveFTPFile(string.Format("{0}|Folder", curDir));
                                        RemoteDirectoryInfo d = session.ListDirectory(curDir);
                                        LoopFiles(curDir, d);
                                    }
                                }
                                else
                                {
                                    SaveFTPFile(string.Format("{0}|{1}|{2}|{3}", dir, fileInfo.Name, fileInfo.Length, fileInfo.LastWriteTime));
                                }
                                break;
                        }                                               
                    }                   
                }

                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e);
                return false;
            }
        }

        public bool GetListOfFileTemp(string dir)
        {
            try
            {
                SessionOptions sessionOptions2 = new SessionOptions
                {
                    Protocol = Protocol.Sftp,
                    HostName = "sftp.allcardtech.com.ph",
                    PortNumber = 2022,
                    UserName = "bayambang",
                    Password = "Argon2016*",
                    SshHostKeyFingerprint = "ssh-ed25519 256 20:c3:ed:d6:8b:ef:83:0c:5f:e0:0b:58:78:93:6b:c6"
                };

                using (Session session = new Session())
                {
                    session.DisableVersionCheck = true;
                    // Connect
                    
                    session.Open(sessionOptions2);


                    //RemoteDirectoryInfo directory =
                    //    session.ListDirectory("/home/martin/public_html");

                    //string parentDir = "/Bayambang/test e/DataCaptureFiles/";
                    string parentDir = config.SftpRemotePath;

                    RemoteDirectoryInfo directory =
                        session.ListDirectory(parentDir);

                    SaveFTPFile(string.Format("{0}|Folder", parentDir));

                    foreach (RemoteFileInfo fileInfo in directory.Files)
                    {
                        switch (fileInfo.Name)
                        {
                            case ".":
                            case "..":
                                break;
                            default:
                                if (fileInfo.IsDirectory)
                                {
                                    string curDir = parentDir + fileInfo.Name;
                                    SaveFTPFile(string.Format("{0}|Folder", curDir));
                                    RemoteDirectoryInfo d = session.ListDirectory(curDir);
                                    LoopFiles(curDir, d);
                                }
                                else
                                {
                                    SaveFTPFile(string.Format("{0}|{1}|{2}|{3}", dir, fileInfo.Name, fileInfo.Length, fileInfo.LastWriteTime));
                                }
                                break;
                        }
                    }
                }

                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e);
                return false;
            }
        }

        public void LoopFiles(string dir, RemoteDirectoryInfo directory)
        {
            foreach (RemoteFileInfo fileInfo in directory.Files)
            {
                switch (fileInfo.Name)
                {
                    case ".":
                    case "..":
                        break;
                    default:
                        SaveFTPFile(string.Format("{0}|{1}|{2}|{3}", dir, fileInfo.Name, fileInfo.Length, fileInfo.LastWriteTime));
                        break;
                }
            }           
        }

        private void SaveFTPFile(string value)
        {
            StreamWriter sw = new StreamWriter("sftpFiles.txt", true);
            sw.WriteLine(DateTime.Now.ToString() + "|" + value);
            sw.Close();
            sw.Dispose();
        }


    }
}
