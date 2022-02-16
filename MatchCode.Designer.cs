
namespace MatchingCode
{
    partial class MatchCode
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MatchCode));
            this.txtLabel1 = new System.Windows.Forms.Label();
            this.txtLabel2 = new System.Windows.Forms.Label();
            this.tbxExcelD = new System.Windows.Forms.TextBox();
            this.tbxExcelC = new System.Windows.Forms.TextBox();
            this.rtbLogs = new System.Windows.Forms.RichTextBox();
            this.btnExe = new System.Windows.Forms.Button();
            this.btnOpen1 = new System.Windows.Forms.Button();
            this.btnOpen2 = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.pBarExecuting = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // txtLabel1
            // 
            this.txtLabel1.AutoSize = true;
            this.txtLabel1.Font = new System.Drawing.Font("Microsoft YaHei UI", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtLabel1.Location = new System.Drawing.Point(12, 9);
            this.txtLabel1.Name = "txtLabel1";
            this.txtLabel1.Size = new System.Drawing.Size(253, 30);
            this.txtLabel1.TabIndex = 0;
            this.txtLabel1.Text = "病例组Excel文件路径：";
            // 
            // txtLabel2
            // 
            this.txtLabel2.AutoSize = true;
            this.txtLabel2.Font = new System.Drawing.Font("Microsoft YaHei UI", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtLabel2.Location = new System.Drawing.Point(12, 49);
            this.txtLabel2.Name = "txtLabel2";
            this.txtLabel2.Size = new System.Drawing.Size(253, 30);
            this.txtLabel2.TabIndex = 1;
            this.txtLabel2.Text = "对照组Excel文件路径：";
            // 
            // tbxExcelD
            // 
            this.tbxExcelD.Location = new System.Drawing.Point(256, 9);
            this.tbxExcelD.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbxExcelD.Name = "tbxExcelD";
            this.tbxExcelD.Size = new System.Drawing.Size(376, 27);
            this.tbxExcelD.TabIndex = 2;
            this.tbxExcelD.TextChanged += new System.EventHandler(this.tbxExcelD_TextChanged);
            // 
            // tbxExcelC
            // 
            this.tbxExcelC.Location = new System.Drawing.Point(256, 49);
            this.tbxExcelC.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbxExcelC.Name = "tbxExcelC";
            this.tbxExcelC.Size = new System.Drawing.Size(376, 27);
            this.tbxExcelC.TabIndex = 3;
            this.tbxExcelC.TextChanged += new System.EventHandler(this.tbxExcelC_TextChanged);
            // 
            // rtbLogs
            // 
            this.rtbLogs.Location = new System.Drawing.Point(12, 82);
            this.rtbLogs.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.rtbLogs.Name = "rtbLogs";
            this.rtbLogs.Size = new System.Drawing.Size(745, 315);
            this.rtbLogs.TabIndex = 4;
            this.rtbLogs.Text = "";
            this.rtbLogs.TextChanged += new System.EventHandler(this.rtbLogs_TextChanged);
            // 
            // btnExe
            // 
            this.btnExe.Font = new System.Drawing.Font("Microsoft YaHei UI", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.btnExe.Location = new System.Drawing.Point(642, 402);
            this.btnExe.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnExe.Name = "btnExe";
            this.btnExe.Size = new System.Drawing.Size(114, 36);
            this.btnExe.TabIndex = 5;
            this.btnExe.Text = "执行";
            this.btnExe.UseVisualStyleBackColor = true;
            this.btnExe.Click += new System.EventHandler(this.btnExe_Click);
            // 
            // btnOpen1
            // 
            this.btnOpen1.Location = new System.Drawing.Point(642, 7);
            this.btnOpen1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnOpen1.Name = "btnOpen1";
            this.btnOpen1.Size = new System.Drawing.Size(94, 29);
            this.btnOpen1.TabIndex = 6;
            this.btnOpen1.Text = "加载";
            this.btnOpen1.UseVisualStyleBackColor = true;
            this.btnOpen1.Click += new System.EventHandler(this.btnOpen1_Click);
            // 
            // btnOpen2
            // 
            this.btnOpen2.Location = new System.Drawing.Point(642, 47);
            this.btnOpen2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnOpen2.Name = "btnOpen2";
            this.btnOpen2.Size = new System.Drawing.Size(94, 29);
            this.btnOpen2.TabIndex = 7;
            this.btnOpen2.Text = "加载";
            this.btnOpen2.UseVisualStyleBackColor = true;
            this.btnOpen2.Click += new System.EventHandler(this.btnOpen2_Click);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // pBarExecuting
            // 
            this.pBarExecuting.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pBarExecuting.Location = new System.Drawing.Point(0, 443);
            this.pBarExecuting.Name = "pBarExecuting";
            this.pBarExecuting.Size = new System.Drawing.Size(767, 29);
            this.pBarExecuting.TabIndex = 8;
            // 
            // MatchCode
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 472);
            this.Controls.Add(this.pBarExecuting);
            this.Controls.Add(this.btnOpen2);
            this.Controls.Add(this.btnOpen1);
            this.Controls.Add(this.btnExe);
            this.Controls.Add(this.rtbLogs);
            this.Controls.Add(this.tbxExcelC);
            this.Controls.Add(this.tbxExcelD);
            this.Controls.Add(this.txtLabel2);
            this.Controls.Add(this.txtLabel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MatchCode";
            this.Text = "MatchingProgram";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MatchCode_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label txtLabel1;
        private System.Windows.Forms.Label txtLabel2;
        private System.Windows.Forms.TextBox tbxExcelD;
        private System.Windows.Forms.TextBox tbxExcelC;
        private System.Windows.Forms.RichTextBox rtbLogs;
        private System.Windows.Forms.Button btnExe;
        private System.Windows.Forms.Button btnOpen1;
        private System.Windows.Forms.Button btnOpen2;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.ProgressBar pBarExecuting;
    }
}

