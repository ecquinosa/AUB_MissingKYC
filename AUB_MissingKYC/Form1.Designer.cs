namespace AUB_MissingKYC
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnGenerate = new System.Windows.Forms.Button();
            this.txtExcelFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnExcelFile = new System.Windows.Forms.Button();
            this.grid = new System.Windows.Forms.DataGridView();
            this.grid2 = new System.Windows.Forms.DataGridView();
            this.lblTotal = new System.Windows.Forms.Label();
            this.btnReport = new System.Windows.Forms.Button();
            this.chkForce = new System.Windows.Forms.CheckBox();
            this.btnMove = new System.Windows.Forms.Button();
            this.btnTransferToSftp = new System.Windows.Forms.Button();
            this.chkSelectAll = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grid2)).BeginInit();
            this.SuspendLayout();
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(12, 612);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(95, 34);
            this.btnGenerate.TabIndex = 0;
            this.btnGenerate.Text = "GENERATE";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // txtExcelFile
            // 
            this.txtExcelFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtExcelFile.Location = new System.Drawing.Point(155, 19);
            this.txtExcelFile.Name = "txtExcelFile";
            this.txtExcelFile.Size = new System.Drawing.Size(632, 23);
            this.txtExcelFile.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "EXCEL FILE";
            // 
            // btnExcelFile
            // 
            this.btnExcelFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExcelFile.Location = new System.Drawing.Point(793, 19);
            this.btnExcelFile.Name = "btnExcelFile";
            this.btnExcelFile.Size = new System.Drawing.Size(84, 23);
            this.btnExcelFile.TabIndex = 6;
            this.btnExcelFile.Text = "BROWSE";
            this.btnExcelFile.UseVisualStyleBackColor = true;
            this.btnExcelFile.Click += new System.EventHandler(this.btnExcelFile_Click);
            // 
            // grid
            // 
            this.grid.AllowUserToAddRows = false;
            this.grid.AllowUserToDeleteRows = false;
            this.grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grid.Location = new System.Drawing.Point(12, 118);
            this.grid.Name = "grid";
            this.grid.ReadOnly = true;
            this.grid.RowHeadersWidth = 51;
            this.grid.RowTemplate.Height = 24;
            this.grid.Size = new System.Drawing.Size(360, 484);
            this.grid.TabIndex = 11;
            this.grid.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.grid_CellMouseDoubleClick);
            // 
            // grid2
            // 
            this.grid2.AllowUserToAddRows = false;
            this.grid2.AllowUserToDeleteRows = false;
            this.grid2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.grid2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grid2.Location = new System.Drawing.Point(378, 118);
            this.grid2.Name = "grid2";
            this.grid2.ReadOnly = true;
            this.grid2.RowHeadersWidth = 51;
            this.grid2.RowTemplate.Height = 24;
            this.grid2.Size = new System.Drawing.Size(500, 484);
            this.grid2.TabIndex = 12;
            // 
            // lblTotal
            // 
            this.lblTotal.AutoSize = true;
            this.lblTotal.Location = new System.Drawing.Point(375, 612);
            this.lblTotal.Name = "lblTotal";
            this.lblTotal.Size = new System.Drawing.Size(74, 17);
            this.lblTotal.TabIndex = 13;
            this.lblTotal.Text = "TOTAL: 0";
            // 
            // btnReport
            // 
            this.btnReport.Location = new System.Drawing.Point(113, 612);
            this.btnReport.Name = "btnReport";
            this.btnReport.Size = new System.Drawing.Size(95, 34);
            this.btnReport.TabIndex = 14;
            this.btnReport.Text = "REPORT";
            this.btnReport.UseVisualStyleBackColor = true;
            this.btnReport.Click += new System.EventHandler(this.btnReport_Click);
            // 
            // chkForce
            // 
            this.chkForce.AutoSize = true;
            this.chkForce.Location = new System.Drawing.Point(12, 57);
            this.chkForce.Name = "chkForce";
            this.chkForce.Size = new System.Drawing.Size(219, 21);
            this.chkForce.TabIndex = 15;
            this.chkForce.Text = "Force generate ManualPack";
            this.chkForce.UseVisualStyleBackColor = true;
            // 
            // btnMove
            // 
            this.btnMove.Font = new System.Drawing.Font("Verdana", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMove.Location = new System.Drawing.Point(12, 652);
            this.btnMove.Name = "btnMove";
            this.btnMove.Size = new System.Drawing.Size(196, 34);
            this.btnMove.TabIndex = 16;
            this.btnMove.Text = "MOVE TO TRANSFER FOLDER AND ZIP";
            this.btnMove.UseVisualStyleBackColor = true;
            this.btnMove.Click += new System.EventHandler(this.btnMove_Click);
            // 
            // btnTransferToSftp
            // 
            this.btnTransferToSftp.Font = new System.Drawing.Font("Verdana", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTransferToSftp.Location = new System.Drawing.Point(214, 652);
            this.btnTransferToSftp.Name = "btnTransferToSftp";
            this.btnTransferToSftp.Size = new System.Drawing.Size(158, 34);
            this.btnTransferToSftp.TabIndex = 17;
            this.btnTransferToSftp.Text = "TRANSFER TO SFTP";
            this.btnTransferToSftp.UseVisualStyleBackColor = true;
            this.btnTransferToSftp.Click += new System.EventHandler(this.btnTransferToSftp_Click);
            // 
            // chkSelectAll
            // 
            this.chkSelectAll.AutoSize = true;
            this.chkSelectAll.Location = new System.Drawing.Point(12, 84);
            this.chkSelectAll.Name = "chkSelectAll";
            this.chkSelectAll.Size = new System.Drawing.Size(93, 21);
            this.chkSelectAll.TabIndex = 18;
            this.chkSelectAll.Text = "Select All";
            this.chkSelectAll.UseVisualStyleBackColor = true;
            this.chkSelectAll.CheckedChanged += new System.EventHandler(this.chkSelectAll_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(889, 707);
            this.Controls.Add(this.chkSelectAll);
            this.Controls.Add(this.btnTransferToSftp);
            this.Controls.Add(this.btnMove);
            this.Controls.Add(this.chkForce);
            this.Controls.Add(this.btnReport);
            this.Controls.Add(this.lblTotal);
            this.Controls.Add(this.grid2);
            this.Controls.Add(this.grid);
            this.Controls.Add(this.btnExcelFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtExcelFile);
            this.Controls.Add(this.btnGenerate);
            this.Font = new System.Drawing.Font("Verdana", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AUB - KYC";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grid2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.TextBox txtExcelFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnExcelFile;
        private System.Windows.Forms.DataGridView grid;
        private System.Windows.Forms.DataGridView grid2;
        private System.Windows.Forms.Label lblTotal;
        private System.Windows.Forms.Button btnReport;
        private System.Windows.Forms.CheckBox chkForce;
        private System.Windows.Forms.Button btnMove;
        private System.Windows.Forms.Button btnTransferToSftp;
        private System.Windows.Forms.CheckBox chkSelectAll;
    }
}

