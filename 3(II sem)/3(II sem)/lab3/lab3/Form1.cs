using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace lab3
{
    public partial class Form1 : Form
    {
        List<Goods> goods = new List<Goods>();
        public Form1()
        {
            InitializeComponent();
            //goods = new Goods().Initialize();
            //dataGridView1.DataSource = goods;
            //int dfd = 2488;
            //label5.Text = "Итого" + dfd.ToString();
            var source = new BindingSource();
            goods = new List<Goods> { new Goods
              {
                Id=1,
                Product="Ананасы",
                Count=3,
                Price=16,
              },
              new Goods
              {
                Id=2,
                Product="Апельсины",
                Count=10,
                Price=40,
              },
              new Goods
              {
                Id=3,
                Product="Яблоки",
                Count=8,
                Price=80,
              },
              new Goods
              {
                Id=4,
                Product="Лимоны",
                Count=35,
                Price=50,
              } };
            source.DataSource = goods;
            dataGridView1.DataSource = source;
            TotalSumm();
        }
        public int TotalSumm()
        {
            int res = 0;
            foreach (Goods good in goods)
            {
                res += good.Sum;
            }
            label5.Text = "Итого: " + res.ToString();
            return res;
        }
        private void exportToWordButton_Click(object sender, EventArgs e)
        {
            // создаем приложение ворд
            Word.Application winword = new Word.Application();

            // добавляем документ
            Word.Document document = winword.Documents.Add();

            // добавляем параграф с номером накладной и выбранной датой
            Word.Paragraph invoicePar = document.Content.Paragraphs.Add();
            DateTime? selectDate = dateTimePicker1.Value;
            if (selectDate != null)
                invoicePar.Range.Text = string.Concat("Расходная накладная № ", textBox3.Text, " от ", selectDate.Value.ToString("dd.MM.yyyy"));
            else
                invoicePar.Range.Text = string.Concat("Расходная накладная № ", textBox3.Text);
            invoicePar.Range.Font.Name = "Times new roman";
            invoicePar.Range.Font.Size = 14;
            invoicePar.Range.InsertParagraphAfter();

            // добавляем параграф с поставщиком
            Word.Paragraph providerPar = document.Content.Paragraphs.Add();
            providerPar.Range.Text = string.Concat("Поставщик: ", textBox1.Text);
            providerPar.Range.Font.Name = "Times new roman";
            providerPar.Range.Font.Size = 14;
            providerPar.Range.InsertParagraphAfter();

            // добавляем параграф с потребителем
            Word.Paragraph customerPar = document.Content.Paragraphs.Add();
            customerPar.Range.Text = "Покупатель: " + textBox2.Text;
            customerPar.Range.Font.Name = "Times new roman";
            customerPar.Range.Font.Size = 14;
            customerPar.Range.InsertParagraphAfter();

            // формируем таблицу
            //int nRows = dataGridView1.Columns.Count;
            Word.Table myTable = document.Tables.Add(customerPar.Range, dataGridView1.Rows.Count, dataGridView1.Columns.Count);
            myTable.Borders.Enable = 1;

            // устанавливаем названия колонок
            var headerRow = myTable.Rows[1].Cells;
            headerRow[1].Range.Text = "№";
            headerRow[2].Range.Text = "Товар";
            headerRow[3].Range.Text = "Количество";
            headerRow[4].Range.Text = "Цена";
            headerRow[5].Range.Text = "Сумма";

            // добавляем данные из таблицы в ворд
            for (int i = 2; i < goods.Count + 2; i++)
            {
                var dataRow = myTable.Rows[i].Cells;
                dataRow[1].Range.Text = goods[i - 2].Id.ToString();
                dataRow[2].Range.Text = goods[i - 2].Product;
                dataRow[3].Range.Text = goods[i - 2].Count.ToString();
                dataRow[4].Range.Text = goods[i - 2].Price.ToString();
                dataRow[5].Range.Text = goods[i - 2].Sum.ToString();
            }

            invoicePar.Range.Text = string.Concat("Итого: ", TotalSumm().ToString());
            invoicePar.Range.Font.Name = "Times new roman";
            invoicePar.Range.Font.Size = 14;
            invoicePar.Range.InsertParagraphAfter();

            // указываем в какой файл сохранить
            //object filename = "Word Files|*.doc";
            Stream MyStream;
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word Files|*.doc";
            var filename = "";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if((MyStream = saveFileDialog.OpenFile()) != null)
                {
                    filename = saveFileDialog.FileName;
                    MyStream.Close();
                    MessageBox.Show("Данные успешно экспортированы в Word!", "Экспорт завершен", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            document.SaveAs(filename, Word.WdSaveFormat.wdFormatDocumentDefault);
            document.Close();
            winword.Quit();
        }
        private void exportToExcelButton_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1] = "Расходная накладная № " + textBox3.Text + " от " + dateTimePicker1.Text;
            worksheet.Cells[2, 1] = "Поставщик: " + textBox1.Text;
            worksheet.Cells[3, 1] = "Покупатель: " + textBox2.Text;

            // Add headers
            worksheet.Cells[5, 1] = "ID";
            worksheet.Cells[5, 2] = "Товар";
            worksheet.Cells[5, 3] = "Количество";
            worksheet.Cells[5, 4] = "Цена";
            worksheet.Cells[5, 5] = "Сумма";

            // Add data to Excel
            for (int i = 0; i < goods.Count; i++)
            {
                worksheet.Cells[i + 6, 1] = goods[i].Id;
                worksheet.Cells[i + 6, 2] = goods[i].Product;
                worksheet.Cells[i + 6, 3] = goods[i].Count;
                worksheet.Cells[i + 6, 4] = goods[i].Price;
                worksheet.Cells[i + 6, 5] = goods[i].Sum;
            }

            worksheet.Cells[goods.Count() + 8, 1] = $"Итого: {TotalSumm().ToString()}";

            Excel.Range tableRange = worksheet.Range[worksheet.Cells[5, 1], worksheet.Cells[goods.Count() + 6, 5]];
            tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tableRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            // Specify the file to save
            //object filename = "Excel Files|*.xlsx";
            Stream MyStream;
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xlsx";
            var filename = "";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if ((MyStream = saveFileDialog.OpenFile()) != null)
                {
                    filename = saveFileDialog.FileName;
                    MyStream.Close();
                    MessageBox.Show("Данные успешно экспортированы в Excel!", "Экспорт завершен", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            workbook.SaveAs(filename);
            workbook.Close();
            excelApp.Quit();
        }
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                goods.Add(new Goods
                {
                    Id = Convert.ToInt32(dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].Value),
                    Product = dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[1].Value.ToString(),
                    Count = Convert.ToInt32(dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[2].Value),
                    Price = Convert.ToInt32(dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[3].Value),
                });
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            label5.Text = "Итого: " + TotalSumm().ToString();
        }
    }
}
