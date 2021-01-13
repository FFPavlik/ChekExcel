using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;
using System.Reflection;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using System.Threading;
using System.Deployment.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace CheckedExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Version version = null;

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion;
            }
            lbInfo.Anchor = AnchorStyles.Bottom;
            lbInfo.Text = "Версия публикации " + version +
                "     Версия приложения " + System.Windows.Forms.Application.ProductVersion.ToString();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        ExcelPackage excelPackage;
        ExcelWorksheet namedWorksheet;
        FileInfo fi;
        List<string> path = new List<string>(); // Пути к файлам
        List<int> errorsLine = new List<int>(); // Строка с ошибкой.

        int rowCount;
        string PathTofile;
        int column_10;
        int column_15;
        int column_16;
        int column_17;
        int column_18;
        int column_20;
        int column_21;
        int column_22;
        
        private void btDialog_Click(object sender, EventArgs e)
        {
            path.Clear();
            tbPath.Clear();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                PathTofile = folderBrowserDialog1.SelectedPath;
            }

            foreach (string filestr in Directory.GetFiles(PathTofile, "*", SearchOption.AllDirectories))
            {
                FileInfo file = new FileInfo(filestr);
                path.Add(file.FullName);
                tbPath.Text += file.Name + "\r\n";

                if (file.Extension.ToLower() != ".xlsx" && file.Extension.ToLower() != ".xls")
                {
                    tbPath.Clear();
                    MessageBox.Show("В папке должны быть только excel файлы", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }

        private void btStart_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            tBinfo.Clear();
            progressBar1.Maximum = path.Count;

            if (!rBizg.Checked && !rBpot.Checked && !rBformat.Checked && !rBimport.Checked)
            {
                MessageBox.Show("Выберите режим", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (tbPath.Text == "")
            {
                MessageBox.Show("Выберите папку с excel файлами", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (rBimport.Checked)
            {
                importPDF(path);
                return;
            }

            for (int l = 0; l < path.Count; l++)
            {
                fi = new FileInfo(path[l]);

                try
                {
                    excelPackage = new ExcelPackage(fi);
                }
                catch (Exception)
                {
                    MessageBox.Show("Ошибка при открытии файла " + fi.Name + " \r\n Ошибку может вызвать: \r\n Незакрытый файл \r\n Файл Excel c форматом .xls", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                if (rBizg.Checked || rBpot.Checked)
                {
                    searchErrors(excelPackage);
                }

                if (rBformat.Checked)
                {
                    defaultFormating(excelPackage);
                }
            }
            MessageBox.Show("Проверка завершена", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            progressBar1.Value = 0;
        }

        void defaultFormating(ExcelPackage excelPackage)
        {
            namedWorksheet = excelPackage.Workbook.Worksheets[0];
            rowCount = namedWorksheet.Dimension.End.Row;

            for (int i = 13; i < rowCount; i++)
            {
                if (namedWorksheet.Cells["A" + i + ":Y" + i].Style.Fill.BackgroundColor.Rgb == "FFFF5033") // Диапазон
                {
                    namedWorksheet.Row(i).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    namedWorksheet.Row(i).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                }
                if (namedWorksheet.Cells[i, 1].Merge)
                {
                    namedWorksheet.Cells["A" + i + ":Y" + i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    namedWorksheet.Cells["A" + i + ":Y" + i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }
            }

            try
            {
                excelPackage.Save();
                tBinfo.AppendText("Проверка файла " + fi.Name + " завершена \r\n");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка при сохранении файла " + fi.Name + " Необходимо закрыть все файлы перед проверкой");
                tBinfo.AppendText("Проверка файла " + fi.Name + " НЕ завершена \r\n");
            }

            tBinfo.AppendText("Проверка файла " + fi.Name + " завершена \r\n");
        }

        void searchErrors(ExcelPackage excelPackage)
        {
            namedWorksheet = excelPackage.Workbook.Worksheets[0];
            rowCount = namedWorksheet.Dimension.End.Row;

            for (int i = 13; i < rowCount; i++)
            {
                try
                {
                    column_10 = Convert.ToInt32(namedWorksheet.Cells[i, 10].Value);
                    column_15 = Convert.ToInt32(namedWorksheet.Cells[i, 15].Value);
                    column_16 = Convert.ToInt32(namedWorksheet.Cells[i, 16].Value);
                    column_17 = Convert.ToInt32(namedWorksheet.Cells[i, 17].Value);
                    column_18 = Convert.ToInt32(namedWorksheet.Cells[i, 18].Value);
                    column_20 = Convert.ToInt32(namedWorksheet.Cells[i, 20].Value);
                    column_21 = Convert.ToInt32(namedWorksheet.Cells[i, 21].Value);
                    column_22 = Convert.ToInt32(namedWorksheet.Cells[i, 22].Value);
                }
                catch (Exception)
                {
                    printErrors(i, 1);
                    continue;
                }

                for (int j = 16; j < 22; j++)
                {
                    if (j == 19)
                        continue;

                    if (Convert.ToInt32(namedWorksheet.Cells[i, j].Value) < 0)   // Общие условие
                    {
                        errorsLine.Add(i);
                        printColorRed(i);
                        printErrors(i, 2);
                    }
                    else checkFixErrors(i);
                }
                
                if (column_15 > column_10)
                {
                    errorsLine.Add(i);
                    printColorRed(i);
                    printErrors(i, 3);
                }
                else checkFixErrors(i);

                if (rBizg.Checked) 
                {
                    if (column_17 != column_15)
                    {
                        errorsLine.Add(i);
                        printColorRed(i);
                        printErrors(i, 4);
                    }
                    else checkFixErrors(i);

                    if ((column_16 + column_17 + column_18) == 0 && column_22 != 0)
                    {
                        errorsLine.Add(i);
                        printColorRed(i);
                        printErrors(i, 5);
                    }
                    else checkFixErrors(i);

                    if (column_17 == column_10 && column_22 != 100)
                    {
                        errorsLine.Add(i);
                        printColorRed(i);
                        printErrors(i, 6);
                    }
                    else checkFixErrors(i);

                    if (column_10 < (column_16 + column_17 + column_18 + column_20))
                    {
                        errorsLine.Add(i);
                        printColorRed(i);
                        printErrors(i, 7);
                    }
                    else checkFixErrors(i);

                    if (column_18 > 0 && (column_22 < 1 || column_22 > 99))
                    {
                        errorsLine.Add(i);
                        printColorRed(i);
                        printErrors(i, 13);
                    }
                    else checkFixErrors(i);
                }

                if (rBpot.Checked) 
                {
                    if (column_16 + column_17 + column_18 + column_20 != column_15)
                    {
                        errorsLine.Add(i);
                        printColorRed(i);
                        printErrors(i, 8);
                    }
                    else checkFixErrors(i);

                    if ((column_16 > 0 || column_17 > 0 || column_18 > 0 || column_20 > 0) && column_22 != 100)
                    {
                        errorsLine.Add(i);
                        printColorRed(i);
                        printErrors(i, 9);
                    }
                    else checkFixErrors(i);

                    if (column_16 == 0 && column_17 == 0 && column_18 == 0 && column_20 == 0 && column_22 != 0)
                    {
                        errorsLine.Add(i);
                        printColorRed(i);
                        printErrors(i, 10);
                    }
                    else checkFixErrors(i);

                    if (column_15 < (column_16 + column_17 + column_18) && column_21 != (column_16 + column_17 + column_18 - column_15))
                    {
                        errorsLine.Add(i);
                        printColorRed(i);
                        printErrors(i, 11);
                    }
                    else checkFixErrors(i);

                    if (column_15 > (column_16 + column_17 + column_18) && column_20 != column_15 - (column_16 + column_17 + column_18))
                    {
                        errorsLine.Add(i);
                        printColorRed(i);
                        printErrors(i, 12);
                    }
                    else checkFixErrors(i);
                }
            }

            progressBar1.Value++;

            try
            {
                excelPackage.Save();
                tBinfo.AppendText("Проверка файла " + fi.Name + " завершена \r\n");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка при сохранении файла " + fi.Name + " Необходимо закрыть все файлы перед проверкой");
                tBinfo.AppendText("Проверка файла " + fi.Name + " НЕ завершена \r\n");
            }
        }

        void checkFixErrors(int i) // Диапазон. Перекрашиваем строки с исправленными ошибками
        {
            if (errorsLine.Count == 0)
                printColorWhite(i);
            if (errorsLine.Count != 0 && namedWorksheet.Cells["A" + i + ":Y" + i].Style.Fill.BackgroundColor.Rgb == "FFFF5033" && i != errorsLine[errorsLine.Count - 1])
            {
                printColorWhite(i);
            }
        }

        void printErrors(int i, int codeErrors)
        {
            string columnErrors = "";
            string messageErrors = "";

            int rowNumber = dataGridView1.Rows.Add();
            dataGridView1.Rows[rowNumber].Cells[0].Value = fi.Name;
            dataGridView1.Rows[rowNumber].Cells[1].Value = i;

            switch (codeErrors)
            {
                // Обще условие
                case 1: columnErrors = ""; messageErrors = "Не удалось преобразовать формат"; break;
                case 2: columnErrors = ""; messageErrors = "Значение меньше 0"; break;
                case 3: columnErrors = " 10, 15"; messageErrors = "Факт. кол. > Кол. дет"; break;
                // Для изготовителя
                case 4: columnErrors = " 15,17"; messageErrors = "Фактически изготовленное кол. не равно сдано"; break;
                case 5: columnErrors = " 16,17,18"; messageErrors = "Процент выполнения должен = 0 %"; break;
                case 6: columnErrors = " 10,17,22"; messageErrors = "Процент выполнения должен = 100%"; break;
                case 7: columnErrors = " 10,17"; messageErrors = "Проверьте корректность данных! условие  ст.10 < (ст.16 + ст.17 + ст.18 + ст.20)"; break;
                case 13: columnErrors = " 22"; messageErrors = "Процент выполнения должен быть в диапазоне 0 до 100"; break;
                // Для потребителя
                case 8: columnErrors = " 16,17,18,20,15"; messageErrors = "Проверьте корректность данных! условие не выполнилось (ст.16 + ст.17 + ст.18 + ст.20) = ст.15"; break;
                case 9: columnErrors = " 16,17,18,20,22"; messageErrors = "Процент выполнения должен = 100%"; break;
                case 10: columnErrors = " 16,17,18,20,22"; messageErrors = "Процент выполнения должен = 0%"; break;
                case 11: columnErrors = " 16,17,18,20,22"; messageErrors = "Проверьте корректность данных! условие не выполнилось ст.15 < (ст.16 + ст.17 + ст.18) и ст._21 == (ст.16 + ст.17 + ст.18 - ст.15)"; break;
                case 12: columnErrors = " 16,17,18,20,22"; messageErrors = "Проверьте корректность данных! условие ст.15 > (ст.16 + ст.17 + ст.18) и ст.20 = ст.15 - (ст.16 + ст.17 + ст.18)"; break;
            }
            dataGridView1.Rows[rowNumber].Cells[2].Value = columnErrors;
            dataGridView1.Rows[rowNumber].Cells[3].Value = messageErrors;
        }

        void printColorRed(int i)
        {
            namedWorksheet.Cells["A" + i + ":Y" + i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid; // Выделяеи строку в пределах таблицы
            namedWorksheet.Cells["A" + i + ":Y" + i].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 80, 51));
        }
        void printColorWhite(int i)
        {
            namedWorksheet.Cells["A" + i + ":Y" + i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            namedWorksheet.Cells["A" + i + ":Y" + i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
        }

        private void button1_Click(object sender, EventArgs e) // Выгрузка ошибок в Excel
        {
            ExcelPackage excelPackage = new ExcelPackage();

            excelPackage.Workbook.Properties.Author = Environment.UserName;
            excelPackage.Workbook.Properties.Title = "Title of Document";
            excelPackage.Workbook.Properties.Subject = "EPPlus demo export data";
            excelPackage.Workbook.Properties.Created = DateTime.Now;
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

            worksheet.Cells[1, 1].Value = "Файл";
            worksheet.Cells[1, 2].Value = "Строка";
            worksheet.Cells[1, 3].Value = "Столбец";
            worksheet.Cells[1, 4].Value = "Сообщение";

            for (int i = 1; i < dataGridView1.Rows.Count + 1; i++)
            {
                worksheet.Cells[i + 1, 1].Value = dataGridView1.Rows[i - 1].Cells[0].Value;
                worksheet.Cells[i + 1, 2].Value = dataGridView1.Rows[i - 1].Cells[1].Value;
                worksheet.Cells[i + 1, 3].Value = dataGridView1.Rows[i - 1].Cells[2].Value;
                worksheet.Cells[i + 1, 4].Value = dataGridView1.Rows[i - 1].Cells[3].Value;
            }

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            saveFileDialog1.FileName = "ErrorsShowExcel";
            saveFileDialog1.Filter = "Excel files(*.xlsx)|*.xlsx|All files(*.*)|*.*";

            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;

            string SavePathTofile = saveFileDialog1.FileName;

            FileInfo fi1 = new FileInfo(SavePathTofile);
            try
            {
                excelPackage.SaveAs(fi1);
                MessageBox.Show("Файл сохранен");
            }
            catch (Exception)
            { MessageBox.Show("Ошибка при сохранении. Необходимо закрыть файл"); }
        }

        void importPDF(List<string> path) // Импорт в PDF
        {
            string PathTofilePdf = null;

            tBinfo.Clear();

            if (folderBrowserDialog2.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            PathTofilePdf = folderBrowserDialog2.SelectedPath;

            for (int l = 0; l < path.Count; l++)
            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook wb = excel.Workbooks.Open(path[l]);
                Worksheet sheet = (Worksheet)excel.Sheets[1];

                try
                {
                    excel.ActivePrinter = "Microsoft XPS Document Writer (Ne00:)";
                }
                catch (Exception)
                {
                    MessageBox.Show("Не удалось выбрать принтер");
                }

                sheet.PageSetup.LeftMargin = 10;
                sheet.PageSetup.RightMargin = 10;
                sheet.PageSetup.TopMargin = 10;
                sheet.PageSetup.BottomMargin = 10;
                sheet.PageSetup.HeaderMargin = 10;
                sheet.PageSetup.FooterMargin = 10;
                sheet.PageSetup.BottomMargin = 10;
                sheet.PageSetup.CenterVertically = false;
                sheet.PageSetup.CenterHorizontally = true;
                sheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                sheet.PageSetup.PaperSize = XlPaperSize.xlPaperA3;
                sheet.PageSetup.Zoom = 130;
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = false;

                try
                {
                    sheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, PathTofilePdf + "/" + wb.Name, Excel.XlFixedFormatQuality.xlQualityStandard);
                    tBinfo.AppendText("Файла " + wb.Name + " конвертирован в PDF \r\n");
                }
                catch (Exception)
                {
                    MessageBox.Show("Ошибка при экспорте файла");
                    tBinfo.AppendText("Ошибка при конвертации файла " + wb.Name + " конвертация не завершена \r\n");
                }
                finally
                {
                    wb.Close(false, PathTofilePdf + wb.Name);
                }
            }

            MessageBox.Show("Конвертация завершена");
        }

        private void rBformat_CheckedChanged(object sender, EventArgs e)
        {
            btStart.Text = "Вернуть исходное форматирование";
        }

        private void rBpot_CheckedChanged(object sender, EventArgs e)
        {
            btStart.Text = "Начать проверку";
        }

        private void rBizg_CheckedChanged(object sender, EventArgs e)
        {
            btStart.Text = "Начать проверку";
        }

        private void rBimport_CheckedChanged(object sender, EventArgs e)
        {
            btStart.Text = "Экспортировать в PDF";
        }

    }
}
