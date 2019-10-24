using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
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
using Microsoft.Win32;
using OfficeOpenXml;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Grid;

namespace PWC_to_Jobboss
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string SelectedDatabase { get; set; }
        public string SelectedInstance { get; set; }

        private List<ExcelLine> _data = null;

        public MainWindow()
        {
            InitializeComponent();

            var culture = new System.Globalization.CultureInfo(ConfigurationManager.AppSettings["CultureToUse"]);
            System.Threading.Thread.CurrentThread.CurrentCulture = culture;
            System.Threading.Thread.CurrentThread.CurrentUICulture = culture;

            if (!System.Diagnostics.Debugger.IsAttached)
            {
                ConnectionSelection dlg = new ConnectionSelection();
                try
                {
                    if (dlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(dlg.SelectedDatabase) && !string.IsNullOrWhiteSpace(dlg.SelectedInstance))
                    {
                        this.SelectedDatabase = dlg.SelectedDatabase;
                        this.SelectedInstance = dlg.SelectedInstance;

                        this.Title = string.Concat("PWC to Jobboss - v1.0.3.0 - ", this.SelectedInstance, " - ", this.SelectedDatabase);
                    }
                    else
                    {
                        this.Close();
                    }
                }
                catch { this.Close(); }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _data = new List<ExcelLine>();
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Fichiers Excel|*.xlsx";
                openFileDialog.Multiselect = false;
                if (openFileDialog.ShowDialog() == true)
                {
                    System.IO.FileInfo fi = new System.IO.FileInfo(openFileDialog.FileName);
                    using (ExcelPackage package = new ExcelPackage(fi))
                    {
                        ExcelWorksheet sheet = package.Workbook.Worksheets[1];

                        string companyName = string.Empty;
                        using (SqlConnection cn = new SqlConnection(string.Concat("Data Source=", this.SelectedInstance, ";Initial Catalog=", this.SelectedDatabase, ";user id=support;password=lonestar;MultipleActiveResultSets=True;")))
                        {
                            cn.Open();

                            for (int i = 2; i <= sheet.Dimension.End.Row; i++)
                            {
                                var line = new ExcelLine
                                {
                                    Name = Util.ToString(sheet.GetValue(i, 1)),
                                    PurchaseDoc = Util.ToString(sheet.GetValue(i, 2)),
                                    Item = Util.ToString(sheet.GetValue(i, 3)),
                                    Material = Util.ToString(sheet.GetValue(i, 4)),
                                    ShortText = Util.ToString(sheet.GetValue(i, 5)),
                                    E = Util.ToString(sheet.GetValue(i, 6)),
                                    NetPrice = Util.ToDecimal(sheet.GetValue(i, 7)),
                                    NetPriceCurrency = Util.ToString(sheet.GetValue(i, 8)),
                                    DocItem = Util.ToDecimal(sheet.GetValue(i, 9)),
                                    DocItemQty = Util.ToString(sheet.GetValue(i, 10)),
                                    OutstQty = Util.ToDecimal(sheet.GetValue(i, 11)),
                                    StateDelDate = Util.ToDateTimeExcel(sheet.GetValue(i, 12)),
                                    DeliveryDate = Util.ToDateTimeExcel(sheet.GetValue(i, 13)),
                                    P = Util.ToString(sheet.GetValue(i, 14)),
                                    Status = Util.ToString(sheet.GetValue(i, 15)),
                                    CompanyName = string.Empty,
                                    LnMeso = string.Empty,
                                    Description = string.Empty,
                                    SO = string.Empty,
                                    Job = string.Empty,
                                    Shipped = string.Empty,
                                    PromisedDate = string.Empty
                                };

                                if (string.IsNullOrWhiteSpace(companyName) && !string.IsNullOrWhiteSpace(line.Name))
                                    companyName = line.Name;

                                line.CompanyName = companyName;

                                string _selectQuery = string.Concat("select SO_Detail.SO_Line, isnull(SO_Detail.Description, '') as Description, SO_Detail.Sales_Order, isnull(SO_Detail.Job, '') as Job, Delivery.Shipped_Date, SO_Detail.Promised_Date from dbo.SO_Header inner join dbo.SO_Detail on SO_Header.Sales_Order = SO_Detail.Sales_Order left join Delivery on Delivery.SO_Detail = SO_Detail.SO_Detail where SO_Detail.Material = @Material and SO_Header.Customer_PO = @PO and SO_Detail.SO_Line = @LinePWC");
                                using (SqlCommand cmd = new SqlCommand(_selectQuery, cn))
                                {
                                    cmd.Parameters.Add("@Material", SqlDbType.VarChar).Value = line.Material;
                                    cmd.Parameters.Add("@PO", SqlDbType.VarChar).Value = line.PurchaseDoc;
                                    cmd.Parameters.Add("@LinePWC", SqlDbType.VarChar).Value = line.Item;
                                    SqlDataReader rs = cmd.ExecuteReader();
                                    if (rs.Read())
                                    {
                                        string lnMeso = Util.ToString(rs["SO_Line"]);
                                        string desc = Util.ToString(rs["Description"]);
                                        string so = Util.ToString(rs["Sales_Order"]);
                                        string job = Util.ToString(rs["Job"]);
                                        DateTime? shipped = Util.ToDateTimeN(rs["Shipped_Date"]);
                                        DateTime? promisedDate = Util.ToDateTimeN(rs["Promised_Date"]);

                                        line.LnMeso = lnMeso;
                                        line.Description = desc;
                                        line.SO = so;
                                        line.Job = job;
                                        line.Shipped = shipped != null && shipped.HasValue ? shipped.Value.ToString("dd-MMM-yyyy") : string.Empty;
                                        line.PromisedDate = promisedDate != null && promisedDate.HasValue ? promisedDate.Value.ToString("dd-MMM-yyyy") : string.Empty;
                                    }
                                }

                                _data.Add(line);
                            }

                            cn.Close();
                        }
                    }
                }
                else
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Concat("Erreur lors de la lecture du fichier. ", ex.ToString()), "Attention", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            FullGrid.ItemsSource = null;
            FullGrid.ItemsSource = _data;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (_data == null || _data.Count == 0)
            {
                MessageBox.Show("Aucune donnée à exporter.", "Attention", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string filename = string.Concat("Output_", DateTime.Now.ToString("yyyyMMddHHmmss"));
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = filename; // Default file name
            dlg.DefaultExt = ".pdf"; // Default file extension
            dlg.Filter = "PDF (.pdf)|*.pdf"; // Filter files by extension

            if (dlg.ShowDialog() != true)
                return;

            filename = dlg.FileName;
            try
            {
                PdfDocument doc = new PdfDocument();
                PdfPage page = doc.Pages.Add();
                PdfGrid pdfGrid = new PdfGrid();

                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("Purchdoc");
                dataTable.Columns.Add("Ln Meso");
                dataTable.Columns.Add("Ln PWC");
                dataTable.Columns.Add("Material");
                dataTable.Columns.Add("Description Meso");
                dataTable.Columns.Add("SO");
                dataTable.Columns.Add("Job");
                dataTable.Columns.Add("Liv PWC");
                dataTable.Columns.Add("Shipped");
                dataTable.Columns.Add("Promised_Date");

                foreach (var line in _data)
                {
                    dataTable.Rows.Add(new object[] {
                                    line.PurchaseDoc, //Purchdoc
                                    line.LnMeso, //Ln Meso
                                    line.Item, //Ln PWC
                                    line.Material, // Material
                                    line.Description, //Description Meso
                                    line.SO, //SO
                                    line.Job, //Job
                                    line.DeliveryDateString, //LivPWC
                                    line.Shipped, //Shipped
                                    line.PromisedDate //Promised_Date
                                });
                }

                pdfGrid.DataSource = dataTable;

                pdfGrid.Columns[0].Width = 50;
                pdfGrid.Columns[1].Width = 35;
                pdfGrid.Columns[2].Width = 35;
                pdfGrid.Columns[3].Width = 50;
                pdfGrid.Columns[4].Width = 75;
                pdfGrid.Columns[5].Width = 40;
                pdfGrid.Columns[6].Width = 55;
                pdfGrid.Columns[7].Width = 50;
                pdfGrid.Columns[8].Width = 50;
                pdfGrid.Columns[9].Width = 60;

                PdfGridLayoutFormat format = new PdfGridLayoutFormat();
                format.Break = PdfLayoutBreakType.FitPage;
                format.PaginateBounds = new RectangleF(0, 0, 400, 800);
                pdfGrid.Draw(page, new PointF(10, 30), format);

                PdfGraphics graphics = page.Graphics;
                PdfFont font = new PdfStandardFont(PdfFontFamily.Helvetica, 14, PdfFontStyle.Bold);
                graphics.DrawString(_data.First().CompanyName, font, PdfBrushes.Red, new PointF(10, 0));

                doc.Save(filename);
                doc.Close(true);

                System.Diagnostics.Process.Start(filename);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Concat("Erreur lors de la génération du PDF. ", ex.ToString()), "Attention", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }
    }
}
