using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Diolom
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ChooseXlsx = new OpenFileDialog();
            ChooseXlsx.Title = "Выбрать Excel файл";
            ChooseXlsx.Filter = "Excel файлы |*.xlsx|Все файлы |*.*";
            ChooseXlsx.ShowDialog();
            textBox1.Text = ChooseXlsx.FileName;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ChooseDocx = new OpenFileDialog();
            ChooseDocx.Title = "Выбрать шаблон выписки";
            ChooseDocx.Filter = "Word файлы |*.docx|Все файлы |*.*";
            ChooseDocx.ShowDialog();
            textBox2.Text = ChooseDocx.FileName;
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog ChooseFolder = new FolderBrowserDialog();
            ChooseFolder.Description = "Выбрать папку сохранения";
            ChooseFolder.ShowDialog();
            textBox3.Text = ChooseFolder.SelectedPath;
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            if ((textBox4.Text == null || textBox4.Text == "") || (textBox1.Text == null || textBox1.Text == "") || (textBox2.Text == null || textBox2.Text == "") || (textBox3.Text == null || textBox3.Text == ""))
            {
                MessageBox.Show("Заполните все поля для корректной работы!", "Предупреждение!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Application ExcelApp = new Excel.Application();
            Excel.Workbook ExcelBook = ExcelApp.Workbooks.Open(textBox1.Text);
            Excel.Worksheet ExcelSheet = ExcelBook.Worksheets[Convert.ToInt32(textBox4.Text)];

            Excel.Application ExcelApp1 = new Excel.Application();
            Excel.Workbook ExcelBook1 = ExcelApp1.Workbooks.Open(textBox6.Text);
            Excel.Worksheet ExcelSheet1 = ExcelBook1.Worksheets[Convert.ToInt32(textBox7.Text)];

            Excel.Application ExcelApp2 = new Excel.Application();
            Excel.Workbook ExcelBook2 = ExcelApp2.Workbooks.Open(textBox8.Text);
            Excel.Worksheet ExcelSheet2 = ExcelBook2.Worksheets[Convert.ToInt32(textBox9.Text)];

            Excel.Application ExcelApp3 = new Excel.Application();
            Excel.Workbook ExcelBook3 = ExcelApp3.Workbooks.Open(textBox10.Text);
            Excel.Worksheet ExcelSheet3 = ExcelBook3.Worksheets[1];

            string Addr = ExcelSheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Address.ToString();
            string Addr1 = ExcelSheet1.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Address.ToString();
            string Addr2 = ExcelSheet2.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Address.ToString();
            string Addr3 = ExcelSheet3.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Address.ToString();

            progressBar1.Maximum = Convert.ToInt32(Addr.Remove(0, Addr.LastIndexOf('$') + 1)) - 4;
            progressBar1.Value = 0;

            Excel.Range ExcelRange = ExcelSheet.Range["B12", Addr];
            Excel.Range ExcelRange1 = ExcelSheet1.Range["E2", Addr1];
            Excel.Range ExcelRange2 = ExcelSheet2.Range["F2", Addr2];
            Excel.Range ExcelRange3 = ExcelSheet3.Range["A1", Addr3];

            string[] PathArray = new string[Convert.ToInt32(Addr.Remove(0, Addr.LastIndexOf('$') + 1)) - 4];
            int PathCount = 0;
            int MarkCount2 = 0;

            foreach (Excel.Range UsedRow in ExcelRange.Rows)
            {
                Word.Application WordApp = new Word.Application();
                Word.Document WordDoc = WordApp.Documents.Open(textBox2.Text);
                WordDoc.Activate();

                Word.Application WordApp1 = new Word.Application();
                Word.Document WordDoc1 = WordApp1.Documents.Open(textBox5.Text);
                WordDoc1.Activate();

                Word.Application WordApp2 = new Word.Application();
                Word.Document WordDoc2 = WordApp2.Documents.Open(textBox11.Text);
                WordDoc2.Activate();

                int MarkCount = 1;
                int MarkCount1 = 1;
                int ColumnIndex = 1;
                MarkCount2 = 1;

                foreach (Excel.Range UsedCell in UsedRow.Cells)
                {
                    string CheckCell = (string)(ExcelSheet.Cells[10, UsedCell.Column] as Excel.Range).Value;

                    if (CheckCell != null)
                    {
                        if (!CheckCell.Contains("курсов"))
                        {
                            try
                            {
                                switch (UsedCell.Value2.ToString())
                                {
                                    case "3":
                                        WordDoc.Bookmarks[MarkCount++].Range.Text = "удовлетворительно";
                                        break;
                                    case "4":
                                        WordDoc.Bookmarks[MarkCount++].Range.Text = "хорошо";
                                        break;
                                    case "5":
                                        WordDoc.Bookmarks[MarkCount++].Range.Text = "отлично";
                                        break;
                                    case "зачет":
                                        WordDoc.Bookmarks[MarkCount++].Range.Text = "зачтено";
                                        break;
                                    default:
                                        break;
                                }
                            }
                            catch { }
                        }
                        else
                        {
                            try
                            {
                                MarkCount2++;
                                switch (UsedCell.Value2.ToString())
                                {
                                    case "3":
                                        WordDoc1.Bookmarks[MarkCount1++].Range.Text = "удовлетворительно";
                                        break;
                                    case "4":
                                        WordDoc1.Bookmarks[MarkCount1++].Range.Text = "хорошо";
                                        break;
                                    case "5":
                                        WordDoc1.Bookmarks[MarkCount1++].Range.Text = "отлично";
                                        break;
                                    default:
                                        break;
                                }
                            }
                            catch { }  
                        }
                    } 
                }

                try
                {
                    Excel.Range Name = UsedRow.Cells[1, ColumnIndex++];
                    string[] NameWords = Name.Value2.ToString().Split(' ');
                    WordDoc.SaveAs2(textBox3.Text + @"\" + Name.Value2.ToString() + ".docx");
                    WordDoc1.SaveAs2(textBox3.Text + @"\" + Name.Value2.ToString() + "(1).docx");
                    WordDoc2.Bookmarks[1].Range.Text = NameWords[0];
                    WordDoc2.Bookmarks[2].Range.Text = NameWords[1];
                    WordDoc2.Bookmarks[3].Range.Text = NameWords[2];
                    WordDoc2.SaveAs2(textBox3.Text + @"\" + Name.Value2.ToString() + "(2).docx");
                    PathArray[PathCount++] = textBox3.Text + @"\" + Name.Value2.ToString() + "(1).docx";
                    progressBar1.Value++;
                    WordDoc = null;
                    WordDoc1 = null;
                    WordDoc2 = null;
                    WordApp.Quit();
                    WordApp1.Quit();
                    WordApp2.Quit();
                }
                catch { }
            }

            PathCount = 0;
            int KursCount = MarkCount2;
            
            foreach (Excel.Range UsedRow1 in ExcelRange1.Rows)
            {
                if (PathArray[PathCount] != null)
                {
                    Word.Application WordApp1 = new Word.Application();
                    Word.Document WordDoc1 = WordApp1.Documents.Open(PathArray[PathCount]);
                    WordDoc1.Activate();

                    string NewPath = PathArray[PathCount];
                    Word.Application WordApp2 = new Word.Application();
                    Word.Document WordDoc2 = WordApp2.Documents.Open(NewPath.Replace('1', '2'));
                    WordDoc2.Activate();
                    PathCount++;

                    string SecondName = null;
                    foreach (Excel.Range UsedCell in UsedRow1.Cells)
                    {
                        if (SecondName == null)
                            SecondName = UsedCell.Text;
                        WordDoc1.Bookmarks[MarkCount2++].Range.Text = UsedCell.Text;
                    }

                    WordDoc1.Bookmarks[WordDoc1.Bookmarks.Count].Range.Text = (string)(ExcelSheet3.Cells[ExcelRange3.Find(SecondName).Row, 9] as Excel.Range).Value;
                    WordDoc2.Bookmarks[WordDoc2.Bookmarks.Count].Range.Text = (string)(ExcelSheet3.Cells[ExcelRange3.Find(SecondName).Row, 9] as Excel.Range).Value;

                    MarkCount2 -= 5;

                    WordDoc1.Save();
                    WordDoc1 = null;
                    WordApp1.Quit();

                    WordDoc2.Save();
                    WordDoc2 = null;
                    WordApp2.Quit();
                }               
            }

            PathCount = 0;
            foreach (Excel.Range UsedRow1 in ExcelRange2.Rows)
            {
                if (PathArray[PathCount] != null)
                {
                    Word.Application WordApp1 = new Word.Application();
                    Word.Document WordDoc1 = WordApp1.Documents.Open(PathArray[PathCount]);
                    WordDoc1.Activate();

                    Word.Application WordApp2 = new Word.Application();
                    Word.Document WordDoc2 = WordApp2.Documents.Open(PathArray[PathCount].Remove(PathArray[PathCount].IndexOf('('), 3));
                    WordDoc2.Activate();

                    PathCount++;
                    int MarkCount3 = WordDoc1.Bookmarks.Count - KursCount + 1;


                    foreach (Excel.Range UsedCell in UsedRow1.Cells)
                    {
                        if (UsedCell.Text != null && UsedCell.Text != "" && MarkCount3 != WordDoc1.Bookmarks.Count)
                            WordDoc1.Bookmarks[MarkCount3++].Range.Text = UsedCell.Text;
                        else if (MarkCount3 == WordDoc1.Bookmarks.Count)
                            WordDoc2.Bookmarks[WordDoc2.Bookmarks.Count].Range.Text = UsedCell.Text;
                    }

                    WordDoc1.Save();
                    WordDoc1 = null;
                    WordApp1.Quit();

                    WordDoc2.Save();
                    WordDoc2 = null;
                    WordApp2.Quit();
                }    
            }

            progressBar1.Value = 0;
            MessageBox.Show($"Автозаполнение дипломов для группы \"{ExcelSheet.Name}\" прошло успешно!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            ExcelBook = null;
            ExcelApp.Quit();
            ExcelBook1 = null;
            ExcelApp1.Quit();
            ExcelBook2 = null;
            ExcelApp2.Quit();
            ExcelBook3 = null;
            ExcelApp3.Quit();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog ChooseDocx = new OpenFileDialog();
            ChooseDocx.Title = "Выбрать файл приложения";
            ChooseDocx.Filter = "Word файлы |*.docx|Все файлы |*.*";
            ChooseDocx.ShowDialog();
            textBox5.Text = ChooseDocx.FileName;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog ChooseXlsx = new OpenFileDialog();
            ChooseXlsx.Title = "Выбрать Excel файл";
            ChooseXlsx.Filter = "Excel файлы |*.xls|Все файлы |*.*";
            ChooseXlsx.ShowDialog();
            textBox6.Text = ChooseXlsx.FileName;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog ChooseXlsx = new OpenFileDialog();
            ChooseXlsx.Title = "Выбрать Excel файл";
            ChooseXlsx.Filter = "Excel файлы |*.xlsx|Все файлы |*.*";
            ChooseXlsx.ShowDialog();
            textBox8.Text = ChooseXlsx.FileName;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog ChooseXlsx = new OpenFileDialog();
            ChooseXlsx.Title = "Выбрать Excel файл";
            ChooseXlsx.Filter = "Excel файлы |*.xls|Все файлы |*.*";
            ChooseXlsx.ShowDialog();
            textBox10.Text = ChooseXlsx.FileName;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            OpenFileDialog ChooseDocx = new OpenFileDialog();
            ChooseDocx.Title = "Выбрать файл титульника";
            ChooseDocx.Filter = "Word файлы |*.docx|Все файлы |*.*";
            ChooseDocx.ShowDialog();
            textBox11.Text = ChooseDocx.FileName;
        }
    }
}
