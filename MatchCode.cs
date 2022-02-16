using System;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Aspose.Cells;
using static System.Net.Mime.MediaTypeNames;

namespace MatchingCode.UI
{
    public partial class MatchCode : Form
    {
        private string exceld = "";
        private string excelc = "";

        public MatchCode()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            InitializeComponent();
            this.rtbLogs.Text =
                @"Noted:
病例组:数据列A=>G，A:PATIENT_ID B:SEX C:AGE
对比组:数据列A=>G，A:PATIENT_ID B:SEX C:AGE D:TEST_NO E:REPORT_ITEM_NAME F:RESULT G:SOURCE
保证两份数据文件中的表头与程序中的标注一致，否则会报错(卡住)";
        }

        private void btnExe_Click(object sender, EventArgs e)
        {
            this.rtbLogs.Text = "";
            backgroundWorker1.RunWorkerAsync();
        }
           
        private static string GetComparedData(DataTable dt, string gender, int age, string dtesxit)
        {
            DataRow[] drs = dt.Select("SEX = '" + gender + "' and AGE = " + age);
            if (drs.Length > 0)
            { 
                DataTable dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                if (dt0.Rows.Count == 0) goto Next;
                int count = dt0.Rows.Count - 1;
                string number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                return number;
            }
            else
            {
                drs = dt.Select("SEX = '" + gender + "' and AGE <= " + age, "AGE DESC");
                if (drs.Length > 0)
                {
                    DataTable dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                    if (dt0.Rows.Count == 0) goto Next2;
                    int count = dt0.Rows.Count - 1;
                    string number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                    return number;
                }
                else
                {
                    drs = dt.Select("SEX = '" + gender + "' and AGE >= " + age, "AGE ASC");
                    {
                        DataTable dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                        if (dt0.Rows.Count == 0) return "";
                        int count = dt0.Rows.Count - 1;
                        string number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                        return number;
                    }
                }
            }

            Next:
            drs = dt.Select("SEX = '" + gender + "' and AGE <= " + age, "AGE DESC");
            if (drs.Length > 0)
            {
                DataTable dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                if (dt0.Rows.Count == 0) goto Next2;
                int count = dt0.Rows.Count - 1;
                string number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                return number;
            }

            Next2:
            drs = dt.Select("SEX = '" + gender + "' and AGE >= " + age, "AGE ASC");
            if (drs.Length > 0)
            {
                DataTable dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                if (dt0.Rows.Count == 0) return "";
                int count = dt0.Rows.Count - 1;
                string number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                return number;
            }

            return "";
        }

        private static DataTable FilterData(DataTable oldTable, string e) 
        {
            DataTable newTable = oldTable.Clone();
            foreach (DataRow dr in oldTable.Rows) 
            {
                if (!e.Contains(dr["PATIENT_ID"].ToString()))
                    newTable.Rows.Add(dr.ItemArray);
            }
            newTable.AcceptChanges();
            return newTable;
        }
        private void tbxExcelD_TextChanged(object sender, EventArgs e)
        {
            if (tbxExcelD.Text != "")
                exceld = tbxExcelD.Text;
        }

        private void tbxExcelC_TextChanged(object sender, EventArgs e)
        {
            if (tbxExcelC.Text != "")
                excelc = tbxExcelC.Text;
        }

        private void btnOpen1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select file";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "xls files (*.*)|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.tbxExcelD.Text = dialog.FileName;
                exceld = dialog.FileName;
            }
        }

        private void btnOpen2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select file";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "xls files (*.*)|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.tbxExcelC.Text = dialog.FileName;
                excelc = dialog.FileName;
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string existeddata = "";
            if (exceld != "" && excelc != "")
            {
                string urlDisease = exceld; //病例组的文件
                string urlCompared = excelc;  //对照组的数据文件
                this.backgroundWorker1.ReportProgress(0, "[" + DateTime.Now.ToString() + "]" + "病例组与对比组文件输入程序中..." + Environment.NewLine);
                Workbook workbook; Workbook workbookC; Cells cellsC; DataTable dtExcelC;
                try
                {
                    workbookC = new Workbook(urlCompared);
                    cellsC = workbookC.Worksheets[0].Cells;//取对应的工作表.第几个工作表
                    dtExcelC = cellsC.ExportDataTableAsString(0, 0, cellsC.MaxDataRow + 1, cellsC.MaxDataColumn + 1, true);//将Excel表格转换成DataTable进行处理
                }
                catch (Exception ex)
                {
                    this.backgroundWorker1.ReportProgress(0, "[" + DateTime.Now.ToString() + "]" + "对比组文件路径不正确。请检查后再执行。" + ex.Message + Environment.NewLine); return;
                }
                try { workbook = new Workbook(urlDisease); }
                catch (Exception ex)
                {
                    this.backgroundWorker1.ReportProgress(0, "[" + DateTime.Now.ToString() + "]" + "病例组文件路径不正确。请检查后再执行。" + ex.Message + Environment.NewLine); return;
                }
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    this.backgroundWorker1.ReportProgress(25, "[" + DateTime.Now.ToString() + "]" + "导入成功开始对比操作..." + Environment.NewLine);
                    Cells cells = sheet.Cells;
                    var styles = cells["A1"].GetStyle();
                    cells["H1"].PutValue("ControlGroup:");
                    cells["I1"].PutValue("PATIENT_ID"); 
                    cells["J1"].PutValue("SEX"); 
                    cells["K1"].PutValue("AGE"); 
                    cells["L1"].PutValue("TEST_NO"); 
                    cells["M1"].PutValue("REPORT_ITEM_NAME");
                    cells["N1"].PutValue("RESULT");
                    cells["O1"].PutValue("SOURCE");

                    for (int i = 2; i < cells.Rows.Count + 1; i++)//从第二行开始处理数据
                    {
                        if (cells["A" + i.ToString()] != null && cells["A" + i.ToString()].Value != null)
                        {
                            if (!String.IsNullOrEmpty(cells["A" + i.ToString()].Value.ToString())) //患者识别号不为空
                            {
                                string comparedNum = "";
                                try
                                {
                                    comparedNum = GetComparedData(dtExcelC, cells["B" + i.ToString()].Value.ToString(), int.Parse(cells["C" + i.ToString()].Value.ToString().Replace("岁", "")), existeddata);
                                }
                                catch (Exception ex)
                                { 
                                    this.backgroundWorker1.ReportProgress(30, "[" + DateTime.Now.ToString() + "]" + "病例组字段不正确，请检查后再执行。" + ex.Message + Environment.NewLine); return; 
                                }
                                if (comparedNum != "")
                                {
                                    existeddata += comparedNum + "-";
                                    EnumerableRowCollection<DataRow> query = from person in dtExcelC.AsEnumerable()
                                                                             where person.Field<string>("PATIENT_ID") == comparedNum
                                                                             select person;
                                    DataTable newDT = query.AsDataView().ToTable();
                                    if (newDT.Rows.Count > 0)
                                    {
                                        try
                                        {
                                            cells["I" + i.ToString()].PutValue(newDT.Rows[0]["PATIENT_ID"].ToString());
                                            cells["J" + i.ToString()].PutValue(newDT.Rows[0]["SEX"].ToString());
                                            cells["K" + i.ToString()].PutValue(newDT.Rows[0]["AGE"].ToString());
                                            cells["L" + i.ToString()].PutValue(newDT.Rows[0]["TEST_NO"].ToString());
                                            cells["M" + i.ToString()].PutValue(newDT.Rows[0]["REPORT_ITEM_NAME"].ToString());
                                            cells["N" + i.ToString()].PutValue(newDT.Rows[0]["RESULT"].ToString());
                                            cells["O" + i.ToString()].PutValue(newDT.Rows[0]["SOURCE"].ToString());
                                            this.backgroundWorker1.ReportProgress(50, "[" + DateTime.Now.ToString() + "]" + "编号：" + cells["A" + i.ToString()].Value.ToString() + "查询到对照组编号：" + newDT.Rows[0]["PATIENT_ID"].ToString() + Environment.NewLine);
                                        }
                                        catch(Exception ex) 
                                        {
                                            this.backgroundWorker1.ReportProgress(30, "[" + DateTime.Now.ToString() + "]" + "对比组字段不正确，请检查后再执行。" + ex.Message + Environment.NewLine); return;
                                        }                
                                    }
                                    else
                                    {
                                        this.backgroundWorker1.ReportProgress(50, "[" + DateTime.Now.ToString() + "]" + "编号：" + cells["A" + i.ToString()].Value.ToString() + "没有合适的匹配。" + Environment.NewLine);
                                    }

                                }
                            }
                            else
                                this.backgroundWorker1.ReportProgress(50, "[" + DateTime.Now.ToString() + "]" + "PATIENT_ID为空。" + Environment.NewLine);
                        }
                    }
                    for (int col = 0; col < cells.MaxColumn; col++)
                    {
                        sheet.AutoFitColumn(col, 0, cells.MaxRow);
                    }
                }
                workbook.Save(urlDisease);
                this.backgroundWorker1.ReportProgress(100, "[" + DateTime.Now.ToString() + "]" + "执行完毕。请检查文件输出。" + Environment.NewLine);
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string message = e.UserState.ToString();
            this.rtbLogs.Text += message;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        private void rtbLogs_TextChanged(object sender, EventArgs e)
        {
            rtbLogs.SelectionStart = rtbLogs.Text.Length;
            rtbLogs.ScrollToCaret();
        }
    }
}
