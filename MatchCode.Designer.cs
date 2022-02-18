
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
            this.cbxParallelRun = new System.Windows.Forms.CheckBox();
            this.cbxOutPutLog = new System.Windows.Forms.CheckBox();
            this.lblCalcType = new System.Windows.Forms.Label();
            this.cmbxCalcType = new System.Windows.Forms.ComboBox();
            this.nudCalcWidth = new System.Windows.Forms.NumericUpDown();
            this.lblPlusMinus = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.nudCalcWidth)).BeginInit();
            this.SuspendLayout();
            // 
            // txtLabel1
            // 
            resources.ApplyResources(this.txtLabel1, "txtLabel1");
            this.txtLabel1.Name = "txtLabel1";
            // 
            // txtLabel2
            // 
            resources.ApplyResources(this.txtLabel2, "txtLabel2");
            this.txtLabel2.Name = "txtLabel2";
            // 
            // tbxExcelD
            // 
            resources.ApplyResources(this.tbxExcelD, "tbxExcelD");
            this.tbxExcelD.Name = "tbxExcelD";
            this.tbxExcelD.TextChanged += new System.EventHandler(this.tbxExcelD_TextChanged);
            // 
            // tbxExcelC
            // 
            resources.ApplyResources(this.tbxExcelC, "tbxExcelC");
            this.tbxExcelC.Name = "tbxExcelC";
            this.tbxExcelC.TextChanged += new System.EventHandler(this.tbxExcelC_TextChanged);
            // 
            // rtbLogs
            // 
            resources.ApplyResources(this.rtbLogs, "rtbLogs");
            this.rtbLogs.Name = "rtbLogs";
            this.rtbLogs.TextChanged += new System.EventHandler(this.rtbLogs_TextChanged);
            // 
            // btnExe
            // 
            resources.ApplyResources(this.btnExe, "btnExe");
            this.btnExe.Name = "btnExe";
            this.btnExe.UseVisualStyleBackColor = true;
            this.btnExe.Click += new System.EventHandler(this.btnExe_Click);
            // 
            // btnOpen1
            // 
            resources.ApplyResources(this.btnOpen1, "btnOpen1");
            this.btnOpen1.Name = "btnOpen1";
            this.btnOpen1.UseVisualStyleBackColor = true;
            this.btnOpen1.Click += new System.EventHandler(this.btnOpen1_Click);
            // 
            // btnOpen2
            // 
            resources.ApplyResources(this.btnOpen2, "btnOpen2");
            this.btnOpen2.Name = "btnOpen2";
            this.btnOpen2.UseVisualStyleBackColor = true;
            this.btnOpen2.Click += new System.EventHandler(this.btnOpen2_Click);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BackgroundWorker1_RunWorkerCompleted);
            // 
            // pBarExecuting
            // 
            resources.ApplyResources(this.pBarExecuting, "pBarExecuting");
            this.pBarExecuting.Name = "pBarExecuting";
            // 
            // cbxParallelRun
            // 
            resources.ApplyResources(this.cbxParallelRun, "cbxParallelRun");
            this.cbxParallelRun.Name = "cbxParallelRun";
            this.cbxParallelRun.UseVisualStyleBackColor = true;
            this.cbxParallelRun.CheckedChanged += new System.EventHandler(this.cbxParallelRun_CheckedChanged);
            // 
            // cbxOutPutLog
            // 
            resources.ApplyResources(this.cbxOutPutLog, "cbxOutPutLog");
            this.cbxOutPutLog.Name = "cbxOutPutLog";
            this.cbxOutPutLog.UseVisualStyleBackColor = true;
            this.cbxOutPutLog.CheckedChanged += new System.EventHandler(this.cbxOutPutLog_CheckedChanged);
            // 
            // lblCalcType
            // 
            resources.ApplyResources(this.lblCalcType, "lblCalcType");
            this.lblCalcType.Name = "lblCalcType";
            // 
            // cmbxCalcType
            // 
            this.cmbxCalcType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbxCalcType.FormattingEnabled = true;
            resources.ApplyResources(this.cmbxCalcType, "cmbxCalcType");
            this.cmbxCalcType.Name = "cmbxCalcType";
            this.cmbxCalcType.SelectedValueChanged += new System.EventHandler(this.cmbxCalcType_SelectedValueChanged);
            // 
            // nudCalcWidth
            // 
            resources.ApplyResources(this.nudCalcWidth, "nudCalcWidth");
            this.nudCalcWidth.Name = "nudCalcWidth";
            this.nudCalcWidth.ValueChanged += new System.EventHandler(this.nudCalcWidth_ValueChanged);
            // 
            // lblPlusMinus
            // 
            resources.ApplyResources(this.lblPlusMinus, "lblPlusMinus");
            this.lblPlusMinus.CausesValidation = false;
            this.lblPlusMinus.Name = "lblPlusMinus";
            // 
            // MatchCode
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblPlusMinus);
            this.Controls.Add(this.nudCalcWidth);
            this.Controls.Add(this.cmbxCalcType);
            this.Controls.Add(this.lblCalcType);
            this.Controls.Add(this.cbxOutPutLog);
            this.Controls.Add(this.cbxParallelRun);
            this.Controls.Add(this.pBarExecuting);
            this.Controls.Add(this.btnOpen2);
            this.Controls.Add(this.btnOpen1);
            this.Controls.Add(this.btnExe);
            this.Controls.Add(this.rtbLogs);
            this.Controls.Add(this.tbxExcelC);
            this.Controls.Add(this.tbxExcelD);
            this.Controls.Add(this.txtLabel2);
            this.Controls.Add(this.txtLabel1);
            this.MaximizeBox = false;
            this.Name = "MatchCode";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MatchCode_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.nudCalcWidth)).EndInit();
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
        private System.Windows.Forms.CheckBox cbxParallelRun;
        private System.Windows.Forms.CheckBox cbxOutPutLog;
        private System.Windows.Forms.Label lblCalcType;
        private System.Windows.Forms.ComboBox cmbxCalcType;
        private System.Windows.Forms.NumericUpDown nudCalcWidth;
        private System.Windows.Forms.Label lblPlusMinus;
    }
}

