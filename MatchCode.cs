using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;

namespace MatchingCode
{
    public partial class MatchCode : Form
    {
        private string exceld = "";
        private string excelc = "";
        private string filepath = "";
        private string existeddata = "";
        private Workbook workbook = null;
        private int progressnumber = 5 * 2000;

        private const string codetype = "2";

        private ParallelOptions options = new ParallelOptions
        {
            MaxDegreeOfParallelism = 20
        };

        public MatchCode()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            InitializeComponent();
            rtbLogs.Text =
                @"Noted:
病例组:数据列A=>G，A:PATIENT_ID B:SEX C:AGE
对比组:数据列A=>G，A:PATIENT_ID B:SEX C:AGE D:TEST_NO E:REPORT_ITEM_NAME F:RESULT G:SOURCE
保证两份数据文件中的表头与程序中的标注一致，否则会报错(卡住)";           
        }

        private void btnExe_Click(object sender, EventArgs e)
        {
            rtbLogs.Text = "";
            backgroundWorker1.RunWorkerAsync();
        }
           
        private string GetComparedData(DataTable dt, string gender, int age, string dtesxit)
        {
            int count; string number; DataTable dt0;
            DataRow[] drs = dt.Select("SEX = '" + gender + "' and AGE = " + age);
            if (codetype == "1")
            {
                if (drs.Length > 0)
                {
                    dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                    if (dt0.Rows.Count == 0)
                    {
                        drs = dt.Select("SEX = '" + gender + "' and AGE < " + age, "AGE");
                        if (drs.Length > 0)
                        {
                            dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                            if (dt0.Rows.Count == 0)
                            {
                                drs = dt.Select("SEX = '" + gender + "' and AGE > " + age, "AGE");
                                {
                                    dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                                    if (dt0.Rows.Count == 0)
                                    {
                                        return "";
                                    }

                                    count = dt0.Rows.Count - 1;
                                    number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                                    return number;
                                }
                            }
                            else
                            {
                                count = dt0.Rows.Count - 1;
                                number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                                return number;
                            }
                        }
                        else
                        {
                            drs = dt.Select("SEX = '" + gender + "' and AGE > " + age, "AGE");
                            {
                                dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                                if (dt0.Rows.Count != 0)
                                {
                                    count = dt0.Rows.Count - 1;
                                    number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                                    return number;
                                }
                            }
                        }
                    }
                    else
                    {
                        count = dt0.Rows.Count - 1;
                        number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                        return number;
                    }
                }
                else
                {
                    drs = dt.Select("SEX = '" + gender + "' and AGE < " + age, "AGE");
                    if (drs.Length > 0)
                    {
                        dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                        if (dt0.Rows.Count == 0)
                        {
                            drs = dt.Select("SEX = '" + gender + "' and AGE > " + age, "AGE");
                            {
                                dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                                if (dt0.Rows.Count != 0)
                                {
                                    count = dt0.Rows.Count - 1;
                                    number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                                    return number;
                                }
                            }
                        }
                        else
                        {
                            count = dt0.Rows.Count - 1;
                            number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                            return number;
                        }
                    }
                    else
                    {
                        drs = dt.Select("SEX = '" + gender + "' and AGE > " + age, "AGE");
                        {
                            dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                            if (dt0.Rows.Count != 0)
                            {
                                count = dt0.Rows.Count - 1;
                                number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                                return number;
                            }
                        }
                    }
                }
            }
            else if (codetype == "2") 
            {
                if (drs.Length > 0)
                {
                    dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                    if (dt0.Rows.Count == 0)
                    {
                        drs = age >= 5
                            ? dt.Select("SEX = '" + gender + "' and AGE > " + (age - 5).ToString() + " and AGE < " + (age + 5).ToString())
                            : dt.Select("SEX = '" + gender + "' and AGE > 0 and AGE < " + (age + 5).ToString());
                        if (drs.Length > 0)
                        {
                            dt0 = FilterData(drs.CopyToDataTable(), dtesxit);
                            if (dt0.Rows.Count > 0)
                            {
                                count = dt0.Rows.Count - 1;
                                number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                                return number;
                            }
                        }
                    }
                    else
                    {
                        count = dt0.Rows.Count - 1;
                        number = dt0.Rows[new Random().Next(0, count)]["PATIENT_ID"].ToString();
                        return number;
                    }
                }
            }
            return "";
        }

        private DataTable FilterData(DataTable oldTable, string e) 
        {
            DataTable newTable = oldTable.Clone();
            foreach (var dr in from DataRow dr in oldTable.Rows
                               where !e.Contains(dr["PATIENT_ID"].ToString())
                               select dr)
            {
                newTable.Rows.Add(dr.ItemArray);
            }
            newTable.AcceptChanges();
            return newTable;
        }
        private void tbxExcelD_TextChanged(object sender, EventArgs e)
        {
            if (tbxExcelD.Text != "")
            {
                exceld = tbxExcelD.Text;
            }
        }

        private void tbxExcelC_TextChanged(object sender, EventArgs e)
        {
            if (tbxExcelC.Text != "")
            {
                excelc = tbxExcelC.Text;
            }
        }

        private void btnOpen1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Title = "选择文件",
                InitialDirectory = ".\\",
                Filter = "xls files (*.*)|*.xlsx"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                tbxExcelD.Text = dialog.FileName;
                exceld = dialog.FileName;             
                filepath = Path.GetDirectoryName(dialog.FileName) + "\\Log-" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
                File.AppendAllText(filepath, Environment.NewLine + "[" + DateTime.Now.ToString() + "]" + "系统初始化开始处理..." + Environment.NewLine);
            }
        }

        private void btnOpen2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Title = "选择文件",
                InitialDirectory = ".\\",
                Filter = "xls files (*.*)|*.xlsx"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                tbxExcelC.Text = dialog.FileName;
                excelc = dialog.FileName;
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        { 
            if (exceld != "" && excelc != "")
            {
                bool IsCompleted = false;
                backgroundWorker1.ReportProgress(0, "[" + DateTime.Now.ToString() + "]" + "[" + 0 + "%]" + "病例组与对比组文件输入程序中（大文件时间较长）..." + Environment.NewLine);
                DataTable dtExcelC;
                try
                {
                    Cells cellsC = new Workbook(excelc).Worksheets[0].Cells;
                    dtExcelC = cellsC.ExportDataTableAsString(0, 0, cellsC.MaxDataRow + 1, cellsC.MaxDataColumn + 1, true);
                }
                catch (Exception ex)
                {
                    backgroundWorker1.ReportProgress(0, "[" + DateTime.Now.ToString() + "]" + "[" + 0 + "%]" + "对比组文件路径格式不正确。请检查后再执行。" + ex.Message + Environment.NewLine); return;
                }
                try
                {
                    workbook = new Workbook(exceld);
                }
                catch (Exception ex)
                {
                    backgroundWorker1.ReportProgress(0, "[" + DateTime.Now.ToString() + "]" + "[" + 0 + "%]" + "病例组文件路径不正确。请检查后再执行。" + ex.Message + Environment.NewLine); return;
                }
                for (int k = 0; k < workbook.Worksheets.Count; k++)
                {
                    IsCompleted = false;
                    Worksheet wb = workbook.Worksheets[k];
                    backgroundWorker1.ReportProgress(5, "[" + DateTime.Now.ToString() + "]" + "[" + 5 + "%]" + "导入成功开始对比操作..." + Environment.NewLine);
                    Cells cells = wb.Cells;
                    SetHeaders(cells);
                    for (int col = 0; col < cells.MaxColumn; col++)
                    {
                        wb.AutoFitColumn(col, 0, cells.MaxRow);
                    }
                    var result = Parallel.For(1, cells.Rows.Count + 1, options ,
               i =>
               {
                   Interlocked.Add(ref progressnumber, int.Parse(Math.Floor(1.0 / double.Parse(cells.Count.ToString()) * 100 * 2000).ToString()));
                   int num = progressnumber / 2000;
                   double number = double.Parse(progressnumber.ToString()) / 2000.0;
                   if (cells["A" + i.ToString()] != null && cells["A" + i.ToString()].Value != null)
                   {
                       if (!string.IsNullOrEmpty(cells["A" + i.ToString()].Value.ToString())) //患者识别号不为空
                       {
                           string comparedNum = "";
                           try
                           {
                               comparedNum = GetComparedData(dtExcelC, cells["B" + i.ToString()].Value.ToString(), int.Parse(cells["C" + i.ToString()].Value.ToString().Replace("岁", "")), existeddata);
                           }
                           catch (Exception ex)
                           {
                               if (i != 1)
                               {
                                   backgroundWorker1.ReportProgress(num, "[" + DateTime.Now.ToString() + "]" + "[" + number + "%]" + "病例组字段数据不正确，请检查后再执行。" + ex.Message + Environment.NewLine);
                               }
                           }
                           if (comparedNum != "")
                           {
                               lock (existeddata)
                               {
                                   existeddata += comparedNum + "-";
                               }
                               EnumerableRowCollection<DataRow> query = from person in dtExcelC.AsEnumerable()
                                                                        where person.Field<string>("PATIENT_ID") == comparedNum
                                                                        select person;
                               DataTable newDT = query.AsDataView().ToTable();
                               if (newDT.Rows.Count > 0)
                               {
                                   try
                                   {
                                       lock (cells)
                                       {
                                           cells["I" + i.ToString()].PutValue(newDT.Rows[0]["PATIENT_ID"].ToString());
                                           cells["J" + i.ToString()].PutValue(newDT.Rows[0]["SEX"].ToString());
                                           cells["K" + i.ToString()].PutValue(newDT.Rows[0]["AGE"].ToString());
                                           cells["L" + i.ToString()].PutValue(newDT.Rows[0]["TEST_NO"].ToString());
                                           cells["M" + i.ToString()].PutValue(newDT.Rows[0]["REPORT_ITEM_NAME"].ToString());
                                           cells["N" + i.ToString()].PutValue(newDT.Rows[0]["RESULT"].ToString());
                                           cells["O" + i.ToString()].PutValue(newDT.Rows[0]["SOURCE"].ToString());
                                       }
                                       backgroundWorker1.ReportProgress(num, "[" + DateTime.Now.ToString() + "]" + "[" + number + "%]" + "编号：" + cells["A" + i.ToString()].Value.ToString() + "查询到对照组编号：" + newDT.Rows[0]["PATIENT_ID"].ToString() + Environment.NewLine);
                                   }
                                   catch (Exception ex)
                                   {
                                       backgroundWorker1.ReportProgress(num, "[" + DateTime.Now.ToString() + "]" + "[" + number + "%]" + "对比组字段不正确，请检查后再执行。" + ex.Message + Environment.NewLine);
                                   }
                               }
                               else
                               {
                                   backgroundWorker1.ReportProgress(num, "[" + DateTime.Now.ToString() + "]" + "[" + number + "%]" + "编号：" + cells["A" + i.ToString()].Value.ToString() + "没有合适的匹配。" + Environment.NewLine);
                               }
                           }
                       }
                       else
                       {
                           backgroundWorker1.ReportProgress(num, "[" + DateTime.Now.ToString() + "]" + "[" + number + "%]" + "PATIENT_ID为空。" + Environment.NewLine);
                       }
                   }
               });
                    if (result.IsCompleted)
                        IsCompleted = true;

                }
                if (!IsCompleted)
                    Thread.Sleep(2000);
                backgroundWorker1.ReportProgress(100, "[" + DateTime.Now.ToString() + "]" + "[" + 100 + "%]" + "执行完毕。请检查文件输出。" + Environment.NewLine);
            }
        }

        private void SetHeaders(Cells cells) 
        {
            Style styles = cells["A1"].GetStyle();
            cells["H1"].PutValue("ControlGroup:");
            cells["H1"].SetStyle(styles);
            cells["I1"].PutValue("PATIENT_ID");
            cells["I1"].SetStyle(styles);
            cells["J1"].PutValue("SEX");
            cells["J1"].SetStyle(styles);
            cells["K1"].PutValue("AGE");
            cells["K1"].SetStyle(styles);
            cells["L1"].PutValue("TEST_NO");
            cells["L1"].SetStyle(styles);
            cells["M1"].PutValue("REPORT_ITEM_NAME");
            cells["M1"].SetStyle(styles);
            cells["N1"].PutValue("RESULT");
            cells["N1"].SetStyle(styles);
            cells["O1"].PutValue("SOURCE");
            cells["O1"].SetStyle(styles);
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string message = e.UserState.ToString();
            rtbLogs.Text += message;
            pBarExecuting.Value = e.ProgressPercentage;
            File.AppendAllText(filepath, message);
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                workbook.Save(exceld);
                File.AppendAllText(filepath, "[" + DateTime.Now.ToString() + "]" + "异步程序处理完成。" + Environment.NewLine);
            }
            catch (Exception)
            {
                File.AppendAllText(filepath, "[" + DateTime.Now.ToString() + "]" + "异步程序处理完成，程序异常已退出。" + Environment.NewLine);
            }
        }

        private void rtbLogs_TextChanged(object sender, EventArgs e)
        {
            rtbLogs.SelectionStart = rtbLogs.Text.Length;
            rtbLogs.ScrollToCaret();
        }

        private void MatchCode_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (workbook != null && exceld != "" && excelc != "")
            {
                try
                {
                    workbook.Save(exceld);
                    File.AppendAllText(filepath, "[" + DateTime.Now.ToString() + "]" + "程序已退出，文件已保存。" + Environment.NewLine);
                }
                catch (Exception)
                {
                    File.AppendAllText(filepath, "[" + DateTime.Now.ToString() + "]" + "程序已退出。" + Environment.NewLine);
                }

            }
            else if (filepath != "")
            {
                File.AppendAllText(filepath, "[" + DateTime.Now.ToString() + "]" + "程序窗口意外关闭，程序已退出。" + Environment.NewLine);
            }
        }
    }
}
