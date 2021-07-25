using System;

namespace AUB_MissingKYC
{
    class Utilities
    {
        public static string TimeStamp()
        {
            return DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt ");
        }

        public static void ShowInfoMessage(string msg, string header)
        {
            System.Windows.Forms.MessageBox.Show(msg, header, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        public static void ShowWarningMessage(string msg, string header)
        {
            System.Windows.Forms.MessageBox.Show(msg, header, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
        }

        public static void ShowErrorMessage(string msg, string header)
        {
            System.Windows.Forms.MessageBox.Show(msg, header, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
        }
    }
}
