using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcelShow
{
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Excel.Application ObjExcel; // интерфейс Application используется для доступа ко всем методам, свойствам и событиям объекта COM
        private Workbook ObjWorkBook;           // определим ссылку на конкретную книгу
        private _Worksheet ObjWorkSheet;        // определим ссылку на конкретный лист
        private Range excelcells;               // ссылка на конкретную ячейку или группу ячеек
        OpenFileDialog openFileDialog1 = new OpenFileDialog();

        private List<appDescription> _list = new List<appDescription>();        // ссылка на пользовательский класс

        public Form1()
        {
            InitializeComponent();
//            button1.Click += button1_Click;
            openFileDialog1.Filter = "Text files(*.xls)|*.xls|(*.xlsx)|*.xlsx|All files(*.*)|*.*";
 
        }

        public void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            string filename = openFileDialog1.FileName.ToString();
            // читаем файл в строку
            /*            string fileText = System.IO.File.ReadAllText(filename, Encoding.GetEncoding(1251));
                        textBox1.Text = fileText;
            */


            ObjExcel = new Microsoft.Office.Interop.Excel.Application();        // Выделение памяти под документ
//            excelApp.Visible = false;
//            ObjWorkBook = ObjExcel.Workbooks.Open(@"D:\firstDoc1.xls", 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            ObjWorkBook = ObjExcel.Workbooks.Open(filename);     // открываем новый документ
            ObjWorkSheet = (_Worksheet)ObjWorkBook.Sheets[1];       // доступ к первому листу документа

            excelcells = ObjWorkSheet.UsedRange;    // доступ к активной области листа (заполненные ячейки)

            int rowCount = excelcells.Rows.Count;   // количество строк
            int colCount = excelcells.Columns.Count;    // количество столбцов

            for (int i = 1; i <= rowCount; i++) // если ячейки в excel не пустые, добавляем их значения в список _list
            {
                for (int j = 1; j <= colCount; j++)
                {
                    if ((excelcells.Cells[i, j] as Range).Value2 != null)
                    {
                        _list.Add(new appDescription(Convert.ToInt32((excelcells.Cells[i, j] as Range).Value2),
                                                       Convert.ToInt32((excelcells.Cells[i, j+1] as Range).Value2),
                                                       Convert.ToInt32((excelcells.Cells[i, j+2] as Range).Value2)));

                        break;
                    }
                        
                }
            }
            textBox1.Text = filename;
//            ShowData();     // после добавления в _list выводим на экран в DataGridView1
//            ObjExcel.Workbooks.Close();     // завершение процесса Excel
        }

        private object[] GetRowData(appDescription record)
        {
            
                List<object> obj = new List<object>();

                try
                {
                    obj.Add(record.Value1);
                }
                catch
                {
                    obj.Add("-");       // выводим тире в DataGridView, если есть пустая ячейка в списке
                }

                obj.Add(record.Value2);
                obj.Add(record.Value3);

                return obj.ToArray();
            
            }

        private void ShowData()     // вывод на экран в DataGridView из документа Excel
        { 
            foreach(appDescription record in _list)
            {
                dataGridView1.Rows.Add(GetRowData(record));
            }

//            _list.ForEach(d => dataGridView1.Rows.Add(GetRowData(d)));
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void Открыть_Click(object sender, EventArgs e)
        {
            ShowData();     // после добавления в _list выводим на экран в DataGridView1
            ObjExcel.Workbooks.Close();     // завершение процесса Excel
        }
    }
    public class appDescription
    {
        public int Value1 { get; set; }
        public int Value2 { get; set; }
        public int Value3 { get; set; }

        public appDescription() { }

        public appDescription(int _value1, int _value2, int _value3)
        {
            Value1 = _value1;
            Value2 = _value2;
            Value3 = _value3;
        }
    }
}