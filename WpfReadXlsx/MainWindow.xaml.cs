using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfReadXlsx
{
    public partial class MainWindow : Window
    {
        public int count;
        public class ModelExcelWebsbor
        {
            public ModelExcelWebsbor(string way, string okpo)
            {
                this.way = way;
                this.okpo = okpo;
            }
            public string way { get; set; }
            public string okpo { get; set; }
        }

        public class ModelExcelASGS
        {
            public ModelExcelASGS(string okpo, string name, string okved, string typ)
            {
                this.okpo = okpo;
                this.name = name;
                this.okved = okved;
                this.typ = typ;
            }
            public string okpo { get; set; }
            public string name { get; set; }
            public string okved { get; set; }
            public string typ { get; set; }

        }       

        public void ReadXlsxWebsbor()
        {
            List<string> specOPeratorDistinct, onlineWebSborDistinct;

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";

            List<ModelExcelWebsbor> listExcel = new List<ModelExcelWebsbor>();
            string[,] array = new string[2, 50];

            if ((ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK))
            {
                Excel.Application ObjWorkExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

                int lastColumn = lastCell.Column;
                int lastRow = lastCell.Row;

                for (int j = 0; j < lastColumn; j++) //по всем колонкам
                    for (int i = 0; i < lastRow; i++) // по всем строкам
                        array[j, i] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString(); //считываем данные

                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit(); // выйти из Excel                    
            }

            for (int i = 0; i < 50; i++)
            {
                listExcel.Add(new ModelExcelWebsbor(array[0, i], array[1, i]));
            }

            var specOperator = from value in listExcel
                               where value.way == "Спецоператор"
                               select value.okpo;

            var onlineWebSbor = from value in listExcel
                                where value.way == "Онлайн"
                                select value.okpo;

            specOPeratorDistinct = new List<string>(specOperator.Distinct());
            onlineWebSborDistinct = new List<string>(onlineWebSbor.Distinct());
        }

        public async void ReadXlsxASGS()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";

            if ((ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK))
            {
                await Task.Run(() =>
                {
                    List<ModelExcelASGS> listExcel = new List<ModelExcelASGS>();
                    string[,] array = new string[4, 30000];

                    Excel.Application ObjWorkExcel = new Excel.Application();
                    Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
                    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                    var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

                    int lastColumn = lastCell.Column;
                    int lastRow = lastCell.Row;

                    for (int j = 0; j < lastColumn; j++) //по всем колонкам
                        for (int i = 0; i < lastRow; i++) // по всем строкам
                            array[j, i] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString(); //считываем данные

                    ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                    ObjWorkExcel.Quit(); // выйти из Excel
                    
                    for (int i = 0; i < 50; i++)
                    {
                        if (array[0, i] != null)
                        {
                            listExcel.Add(new ModelExcelASGS(array[0, i], array[1, i], array[2, i], array[3, i]));
                           
                            Dispatcher.Invoke(() =>
                            {                            
                                LblCount.Content = ++count;
                            });
                        }
                        else
                        {
                            break;
                        }
                    }                
                });
            }
        }

        public MainWindow()
        {
            InitializeComponent();           
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ReadXlsxWebsbor();
        }

        private void Load_ASGS_Click(object sender, RoutedEventArgs e)
        {
            ReadXlsxASGS();
        }
    }
}
