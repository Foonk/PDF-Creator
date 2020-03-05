using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net;
using Microsoft.Office.Interop.Excel;
using HtmlAgilityPack;
using iText.Html2pdf;
using iText.Kernel.Pdf;
using iText.Layout;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace PDF_Creator
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private string excelFile;
        private string htmlFile;
        private string filesPath;

        private string[] passwords;

        //Параметры
        IList<ParamItem> paramsList = new List<ParamItem>();

        public enum ParamDataType
        {
            Строка,
            Целое_число,
            Дробное_число,
            Дата,
            Время
        }

        

        //Выбор файла Excel
        private void excelFileSelectBtn_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "xls files (*.xlsx;*.xls)|*.xlsx;*.xls";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                excelFile = openFileDialog1.FileName;
                excelFileSelectLabel.Text = excelFile;
            }
        }

        //Выбор файла html с письмом
        private void htmlFileSelectBtn_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "html files (*.html)|*.html";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                htmlFile = openFileDialog1.FileName;
                htmlFileSelectLabel.Text = htmlFile;
            }
        }

        //Выбор пути к файлам для сохранения
        private void folderWithFilesSelectBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                filesPath = folderBrowserDialog1.SelectedPath;
                filesPathLabel.Text = filesPath;
            }
        }

        private void LockInterface(bool lockInt)
        {
            if (lockInt)
            {
                excelFileSelectBtn.Enabled = false;
                htmlFileSelectBtn.Enabled = false;
                loadDataFromFileBtn.Enabled = false;
                folderWithFilesSelectBtn.Enabled = false;
                setPassword.Enabled = false;
                passRowNumber.Enabled = false;
                newParamTitle.Enabled = false;
                newParamRowNumber.Enabled = false;
                newParamCSSclassName.Enabled = false;
                newParamDataType.Enabled = false;
                addParamBtn.Enabled = false;
                deleteParamsBtn.Enabled = false;
                saveConfigBtn.Enabled = false;
                loadConfigBtn.Enabled = false;
            }
            else
            {
                excelFileSelectBtn.Enabled = true;
                htmlFileSelectBtn.Enabled = true;
                loadDataFromFileBtn.Enabled = true;
                folderWithFilesSelectBtn.Enabled = true;
                setPassword.Enabled = true;
                passRowNumber.Enabled = true;
                newParamTitle.Enabled = true;
                newParamRowNumber.Enabled = true;
                newParamCSSclassName.Enabled = true;
                newParamDataType.Enabled = true;
                addParamBtn.Enabled = true;
                deleteParamsBtn.Enabled = true;
                saveConfigBtn.Enabled = true;
                loadConfigBtn.Enabled = true;
            }
        }

        private async void loadDataFromFileBtn_Click(object sender, EventArgs e)
        {
            LockInterface(true);
            try
            {
                //Проверяем заполнение полей
                string errorText = "";
                bool goNext = true;
                if (excelFile == null)
                {
                    goNext = false;
                    errorText += "Не выбран файл Excel\n";
                }
                if (htmlFile == null)
                {
                    goNext = false;
                    errorText += "Не выбран файл html\n";
                }
                if (filesPath == null)
                {
                    goNext = false;
                    errorText += "Не выбрана папка для сохранения\n";
                }
                if (setPassword.Checked && passRowNumber.Text == string.Empty)
                {
                    goNext = false;
                    errorText += "Не указан номер столбца пароля\n";
                }
                if(paramsList.Count == 0)
                {
                    goNext = false;
                    errorText += "Не указаны параметры\n";
                }

                if (goNext)
                {
                    //Создаём приложение.
                    Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                    //Открываем книгу.
                    Workbook ObjWorkBook = ObjExcel.Workbooks.Open(excelFile, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //Выбираем таблицу(лист).
                    Worksheet ObjWorkSheet;
                    ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];


                    for (int i = 0; i < paramsList.Count; i++)
                    {
                        Range paramColumn = ObjWorkSheet.UsedRange.Columns[paramsList[i].RowNumber];
                        Array paramValues = (Array)paramColumn.Cells.Value;
                        paramsList[i].ItemValue = paramValues.OfType<object>().Select(o => o.ToString()).ToArray();
                    }

                    //Получаем пароли из указанного столбца
                    if (setPassword.Checked)
                    {
                        if (passRowNumber.Text != string.Empty)
                        {
                            Range passportNumberColumn = ObjWorkSheet.UsedRange.Columns[Convert.ToInt32(passRowNumber.Text)];
                            Array passportNumberValues = (Array) passportNumberColumn.Cells.Value;
                            passwords = passportNumberValues.OfType<object>().Select(o => o.ToString()).ToArray();
                        }
                    }

                    //Удаляем приложение (выходим из экселя) - ато будет висеть в процессах!
                    ObjWorkBook.Close();
                    ObjExcel.Quit();
                    killExcel();

                    //Для всех значений, полученных из Excel файла формируем временный html
                    for (int i = 1; i < paramsList[0].ItemValue.Length; i++)
                    {
                        //Открываем шаблон HTML в правильной кодировке
                        var doc = new HtmlAgilityPack.HtmlDocument();
                        StreamReader reader = new StreamReader(WebRequest.Create(htmlFile).GetResponse().GetResponseStream(), Encoding.UTF8);
                        doc.Load(reader);

                        for (int k = 0; k < paramsList.Count; k++)
                        {
                            //Ищем ноды с параметром и меняем в них значения
                            HtmlNodeCollection nodeValues = doc.DocumentNode.SelectNodes("//span[contains(@class, '"+ paramsList[k].CSSClassName + "')]");
                            for (int j = 0; j < nodeValues.Count; j++)
                            {
                                //В зависимости от типа данных
                                if (paramsList[k].DataType == ParamDataType.Строка)
                                {
                                    nodeValues[j].InnerHtml = paramsList[k].ItemValue[i];
                                }
                                else if (paramsList[k].DataType == ParamDataType.Целое_число)
                                {
                                    int tempVal = Convert.ToInt32(paramsList[k].ItemValue[i]);
                                    nodeValues[j].InnerHtml = tempVal.ToString();
                                }
                                else if (paramsList[k].DataType == ParamDataType.Дробное_число)
                                {
                                    double tempVal = Convert.ToDouble(paramsList[k].ItemValue[i]);
                                    nodeValues[j].InnerHtml = tempVal.ToString();
                                }
                                else if (paramsList[k].DataType == ParamDataType.Дата)
                                {
                                    DateTime tempVal = Convert.ToDateTime(paramsList[k].ItemValue[i]);
                                    nodeValues[j].InnerHtml = tempVal.ToShortDateString();
                                }
                                else if (paramsList[k].DataType == ParamDataType.Время)
                                {
                                    DateTime tempVal = Convert.ToDateTime(paramsList[k].ItemValue[i]);
                                    nodeValues[j].InnerHtml = tempVal.ToShortTimeString();
                                }
                            }
                        }
                        

                        //Сохраняем html в новом файле
                        FileStream sw = new FileStream(filesPath + "\\" + paramsList[0].ItemValue[i] + ".html", FileMode.Create);
                        doc.Save(sw, Encoding.UTF8);
                        sw.Dispose();

                        reader.Dispose();

                        CreatePdf(i);

                        //Прогресс
                        progressLabel.Text = i + "/" + (paramsList[0].ItemValue.Length - 1);

                        progressBar.Maximum = paramsList[0].ItemValue.Length - 1;
                        progressBar.Step = 1;


                        var progress = new Progress<int>(v =>
                        {
                            progressBar.Value = i;
                        });
                        await Task.Run(() => DoWork(progress));
                    }

                    LockInterface(false);
                    var result = MessageBox.Show("Готово", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    LockInterface(false);
                    var result = MessageBox.Show(errorText, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (System.Exception exception)
            {
                LockInterface(false);
                var result = MessageBox.Show(exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private void CreatePdf(int index)
        {
            try
            {
                FileStream htmlSource = File.Open(filesPath + "\\" + paramsList[0].ItemValue[index] + ".html", FileMode.Open);
                FileStream pdfDest = File.Open(filesPath + "\\" + paramsList[0].ItemValue[index] + ".pdf", FileMode.OpenOrCreate);
                ConverterProperties converterProperties = new ConverterProperties();
                HtmlConverter.ConvertToPdf(htmlSource, pdfDest, converterProperties);
                htmlSource.Dispose();
                pdfDest.Dispose();
                
                //Паролим файл
                if (setPassword.Checked)
                {
                    EncryptPdf(index);
                }

                //Удаляем HTML временный файл
                string fn = filesPath + "\\" + paramsList[0].ItemValue[index] + ".html";
                File.Delete(@"" + fn);
            }
            catch (Exception e)
            {
                var result = MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        public void EncryptPdf(int index)
        {
            Stream input = new FileStream(filesPath + "\\" + paramsList[0].ItemValue[index] + ".pdf", FileMode.Open, FileAccess.Read, FileShare.Read);
            iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(input);
            input.Dispose();
            
            Stream output = new FileStream(filesPath + "\\" + paramsList[0].ItemValue[index] + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            iTextSharp.text.pdf.PdfEncryptor.Encrypt(reader, output, true, passwords[index], passwords[index], iTextSharp.text.pdf.PdfWriter.ALLOW_PRINTING);
            output.Dispose();
            reader.Dispose();
        }

        private void killExcel()
        {
            System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {
                if (PK.MainWindowTitle.Length == 0)
                {
                    PK.Kill();
                }
            }
        }

        private void Calculate(int i)
        {
            double pow = Math.Pow(i, i);
        }

        //Рассчет прогресса
        public void DoWork(IProgress<int> progress)
        {
            for (int j = 0; j < 100000; j++)
            {
                Calculate(j);
                if (progress != null)
                {
                    progress.Report((j + 1) * 100 / 100000);
                }
            }
        }

        private void wordToHtmlLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://wordtohtml.net/");
        }


        

        //Кнопка добавить параметры
        private void addParamBtn_Click(object sender, EventArgs e)
        {
            //Проверяем заполнение полей
            string errorText = "";
            bool goNext = true;
            if (newParamTitle.Text == string.Empty)
            {
                goNext = false;
                errorText += "Не указано название параметра\n";
            }
            if (newParamRowNumber.Text == string.Empty)
            {
                goNext = false;
                errorText += "Не указан номер столбца\n";
            }
            if (newParamCSSclassName.Text == string.Empty)
            {
                goNext = false;
                errorText += "Не указано имя CSS класса\n";
            }
            if (newParamDataType.SelectedItem == null)
            {
                goNext = false;
                errorText += "Не выбран тип данных\n";
            }


            if (goNext)
            {
                //Тип данных
                ParamDataType dt = ParamDataType.Строка;
                string dataTypeString = "";
                if (newParamDataType.SelectedItem == "Строка")
                {
                    dt = ParamDataType.Строка;
                    dataTypeString = "Строка";
                }
                else if (newParamDataType.SelectedItem == "Целое число")
                {
                    dt = ParamDataType.Целое_число;
                    dataTypeString = "Целое число";
                }
                else if (newParamDataType.SelectedItem == "Дробное число")
                {
                    dt = ParamDataType.Дробное_число;
                    dataTypeString = "Дробное число";
                }
                else if (newParamDataType.SelectedItem == "Дата")
                {
                    dt = ParamDataType.Дата;
                    dataTypeString = "Дата";
                }
                else if (newParamDataType.SelectedItem == "Время")
                {
                    dt = ParamDataType.Время;
                    dataTypeString = "Время";
                }

                paramsList.Add(new ParamItem
                {
                    Title = newParamTitle.Text, RowNumber = Convert.ToInt32(newParamRowNumber.Text), CSSClassName = newParamCSSclassName.Text, DataType = dt
                });

                parametersLabel.Text += newParamTitle.Text + " | " + newParamRowNumber.Text + " | " + newParamCSSclassName.Text + " | " + dataTypeString + "\n";

                newParamTitle.Text = string.Empty;
                newParamRowNumber.Text = string.Empty;
                newParamCSSclassName.Text = string.Empty;
            }
            else
            {
                var result = MessageBox.Show(errorText, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Кнопка удаления параметров
        private void deleteParamsBtn_Click(object sender, EventArgs e)
        {
            paramsList = new List<ParamItem>();
            parametersLabel.Text = string.Empty;
        }

        //Кнопка сохранить
        private void saveConfigBtn_Click(object sender, EventArgs e)
        {
            bool goNext = true;
            string errorText = "";
            if (paramsList.Count == 0)
            {
                goNext = false;
                errorText += "Не указаны параметры\n";
            }

            if (goNext)
            {
                JArray toSave = (JArray)JToken.FromObject(paramsList);

                saveFileDialog1.Filter = "json files (*.json)|*.json";
                DialogResult result = saveFileDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    //File.WriteAllText(@"" + saveFileDialog1.FileName, toSave.ToString());

                    using (StreamWriter file = File.CreateText(@"" + saveFileDialog1.FileName))
                    using (JsonTextWriter writer = new JsonTextWriter(file))
                    {
                        toSave.WriteTo(writer);
                    }
                }
            }
            else
            {
                var result = MessageBox.Show(errorText, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Кнопка Загрузить
        private void loadConfigBtn_Click(object sender, EventArgs e)
        {
            paramsList = new List<ParamItem>();
            parametersLabel.Text = string.Empty;

            openFileDialog1.Filter = "json files (*.json)|*.json";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                try
                {
                    using (StreamReader file = File.OpenText(@"" + openFileDialog1.FileName))
                    using (JsonTextReader reader = new JsonTextReader(file))
                    {
                        JArray fromFile1 = (JArray)JToken.ReadFrom(reader);
                        foreach (JObject jo in fromFile1)
                        {
                            ParamDataType dt = ParamDataType.Строка;
                            string dataTypeString = "";
                            if ((int)jo["DataType"] == 0)
                            {
                                dt = ParamDataType.Строка;
                                dataTypeString = "Строка";
                            }
                            if ((int)jo["DataType"] == 1)
                            {
                                dt = ParamDataType.Целое_число;
                                dataTypeString = "Целое число";
                            }
                            if ((int)jo["DataType"] == 2)
                            {
                                dt = ParamDataType.Дробное_число;
                                dataTypeString = "Дробное число";
                            }
                            if ((int)jo["DataType"] == 3)
                            {
                                dt = ParamDataType.Дата;
                                dataTypeString = "Дата";
                            }
                            if ((int)jo["DataType"] == 4)
                            {
                                dt = ParamDataType.Время;
                                dataTypeString = "Время";
                            }

                            paramsList.Add(new ParamItem
                            {
                                Title = (string)jo["Title"],
                                RowNumber = (int)jo["RowNumber"],
                                CSSClassName = (string)jo["CSSClassName"],
                                DataType = dt
                            });

                            parametersLabel.Text += (string)jo["Title"] + " | " + (int)jo["RowNumber"] + " | " + (string)jo["CSSClassName"] + " | " + dataTypeString + "\n";
                        }
                        /*
                        paramsList = ((JArray)fromFile1).Select(x => new ParamItem
                        {
                            Title = (string)x["Title"],
                            RowNumber = (int)x["RowNumber"],
                            CSSClassName = (string)x["CSSClassName"],
                            DataType = ParamDataType.Строка
                            //DataType = (ParamDataType)x["DataType"]
                            //DataType = (ParamDataType DataType)x["DataType"]
                        }).ToList();
                        */
                        //var result1 = MessageBox.Show(paramsList[0].Title.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception exception)
                {
                    var result1 = MessageBox.Show(exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    throw;
                }
            }
        }
    }

    //Класс параметров
    public class ParamItem
    {
        /// <summary>
        /// Название
        /// </summary>
        public string Title;
        /// <summary>
        /// Номер столбца
        /// </summary>
        public int RowNumber;
        /// <summary>
        /// Имя CSS класса
        /// </summary>
        public string CSSClassName;
        /// <summary>
        /// Значение поля
        /// </summary>
        public string[] ItemValue;
        /// <summary>
        /// Тип поля
        /// </summary>
        public MainForm.ParamDataType DataType;
    }
}
