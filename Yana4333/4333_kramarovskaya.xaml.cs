using Microsoft.Win32;
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
using System.Windows.Media.Media3D;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;


namespace Yana4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_kramarovskaya.xaml
    /// </summary>
    public partial class _4333_kramarovskaya : Window
    {
        public _4333_kramarovskaya()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (ofd.ShowDialog() != true)
                return;

            Excel.Application ObjWorkExcel = null;
            Excel.Workbook ObjWorkBook = null;
            Excel.Worksheet ObjWorkSheet = null;

            ObjWorkExcel = new Excel.Application();
            ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            string[,] list = new string[_rows, _columns];

            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (ISRPO31Entities dbContext = new ISRPO31Entities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    if (string.IsNullOrWhiteSpace(list[i, 0]))
                        continue;

                    DateTime? creationDate = null;
                    DateTime? orderTime = null;
                    DateTime? closingDate = null;

                    if (!string.IsNullOrWhiteSpace(list[i, 2]))
                    {
                        if (DateTime.TryParse(list[i, 2], out DateTime parsedCreation))
                        {
                            creationDate = parsedCreation;
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(list[i, 3]))
                    {
                        if (DateTime.TryParse(list[i, 3], out DateTime parsedTime))
                        {
                            orderTime = parsedTime;
                        }
                    }


                    if (!string.IsNullOrWhiteSpace(list[i, 7]))
                    {
                        if (DateTime.TryParse(list[i, 7], out DateTime parsedClosing))
                        {
                            closingDate = parsedClosing;
                        }
                    }

                    dbContext.Orders.Add(new Orders()
                    {
                        Id = int.Parse(list[i, 0]),  
                        OrderCode = list[i, 1],
                        CreationDate = creationDate,
                        OrderTime = orderTime,
                        ClientCode = list[i, 4],
                        Services = list[i, 5],
                        Status = list[i, 6],
                        ClosingDate = closingDate,
                        RentalTime = list[i, 8]
                    });
                }

                dbContext.SaveChanges();
            }

            MessageBox.Show("Данные успешно импортированы!", "Импорт",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {

                List<Orders> allOrders;
                using (ISRPO31Entities dbContext = new ISRPO31Entities())
                {
                    allOrders = dbContext.Orders.ToList();
                }

                var ordersByDate = allOrders
                    .GroupBy(o => o.CreationDate)
                    .OrderBy(g => g.Key)
                    .ToList();

                Excel.Application app = new Excel.Application();
                app.SheetsInNewWorkbook = ordersByDate.Count;
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

                for (int i = 0; i < ordersByDate.Count; i++)
                {
                    Excel.Worksheet worksheet = (Excel.Worksheet)app.Worksheets.Item[i + 1];
                    var dateGroup = ordersByDate[i];

                    string sheetName = dateGroup.Key.HasValue
                        ? dateGroup.Key.Value.ToString("dd-MM-yy").Replace(".", "-")
                        : "NoDate";
                    if (sheetName.Length > 31)
                        sheetName = sheetName.Substring(0, 31);
                    worksheet.Name = sheetName;

                    int startRowIndex = 1;

                    worksheet.Cells[1, 1] = "Id";
                    worksheet.Cells[1, 2] = "Код заказа";
                    worksheet.Cells[1, 3] = "Код клиента";
                    worksheet.Cells[1, 4] = "Услуги";

                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                    headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    headerRange.Font.Italic = true;
                    startRowIndex++;

                    startRowIndex = 2;

                    foreach (var order in dateGroup)
                    {
                        worksheet.Cells[startRowIndex, 1] = order.Id;
                        worksheet.Cells[startRowIndex, 2] = order.OrderCode;
                        worksheet.Cells[startRowIndex, 3] = order.ClientCode;
                        worksheet.Cells[startRowIndex, 4] = order.Services;

                        Excel.Range dataRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex, 4]];
                        dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        startRowIndex++;
                    }

                    worksheet.Cells[startRowIndex, 1] = "Всего заказов:";
                    worksheet.Cells[startRowIndex, 1].Font.Bold = true;
                    worksheet.Cells[startRowIndex, 2] = dateGroup.Count();

                    Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[startRowIndex, 4]];
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                    worksheet.Columns.AutoFit();
                }

                app.Visible = true;

                MessageBox.Show("Данные успешно экспортированы в Excel!", "Экспорт",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}

