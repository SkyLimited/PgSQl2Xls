using Npgsql;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
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

namespace PgSQl2Xls
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string ConnectionString { get; set; }

        public string SQLQuery { get; set; }

        public string OutputFile { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            ConnectionString = "Host=localhost;Port=5432;Username=;Password=;Database=";
            OutputFile = @"C:\Temp\Result.xlsx";
            DataContext = this;
        }


        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }


        private void DoUpload(object sender, RoutedEventArgs e)
        {
            
            using(NpgsqlConnection cs = new NpgsqlConnection(ConnectionString))
            {
                try
                {
                    cs.Open();
                }
                catch(Exception x)
                {
                    MessageBox.Show(x.ToString(), "Connection Failed!");
                }
                using (NpgsqlCommand cmd = new NpgsqlCommand(SQLQuery, cs))
                {
                    cmd.CommandTimeout = 0;
                    int rowCount = 1;
                    var cset = false;


                    var tfs = new FileStream(OutputFile, FileMode.Create);

                    using (ExcelPackage excel = new ExcelPackage())
                    {
                        var sheet = excel.Workbook.Worksheets.Add("Лист 1");
                        try
                        {
                            using (var rd = cmd.ExecuteReader())
                            {
                                while (rd.Read())
                                {


                                    if (!cset)
                                    {
                                        for (int i = 0; i < rd.GetColumnSchema().Count; i++)
                                        {
                                            var c = rd.GetColumnSchema()[i];
                                            sheet.Cells[$"{GetExcelColumnName(i + 1)}{rowCount}"].Value = c.ColumnName;
                                            sheet.Cells[$"{GetExcelColumnName(i + 1)}{rowCount}"].Style.Font.Bold = true;
                                        }
                                        rowCount++;
                                        cset = true;
                                    }
                                    for (int i = 0; i < rd.GetColumnSchema().Count; i++)
                                    {
                                        if (!(rd[i] is DBNull))
                                        {
                                            var xx = Convert.ToString(rd[i]);
                                            if (xx.Length <= 4000)
                                            {

                                                sheet.Cells[$"{GetExcelColumnName(i + 1)}{rowCount}"].Value = rd[i];

                                            }
                                            else
                                            {
                                                sheet.Cells[$"{GetExcelColumnName(i + 1)}{rowCount}"].Value = xx.Substring(0, 3999);
                                            }
                                            if (rd[i] is DateTime)
                                                sheet.Cells[$"{GetExcelColumnName(i + 1)}{rowCount}"].Style.Numberformat.Format = "dd.mm.yyyy";
                                        }
                                    }
                                    rowCount++;

                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error in SQL: {ex.Message}");
                          
                            return;
                        }
                        excel.SaveAs(tfs);
                        tfs.Dispose();
                        MessageBox.Show("Done");

                    }
                }

            }

        }
    }
}
