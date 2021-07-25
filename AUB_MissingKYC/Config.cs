using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AUB_MissingKYC
{
    class Config
    {
        public string KycDumpfileFolder { get; set; }
        public string BankDoneFolder { get; set; }
        public string BankForTransferFolder { get; set; }
        public string DbaseConStr { get; set; }
        public DateTime ReportStartDate { get; set; }

        public string SftpRemotePath { get; set; }
        public string SftpHost { get; set; }
        public string SftpUser { get; set; }
        public string SftpPass { get; set; }
        public string SftpPort { get; set; }
        public string SftpSshHostKeyFingerprint { get; set; }     

    }
}
