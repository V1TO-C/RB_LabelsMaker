using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Windows;


namespace RB_LabelsMaker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click_save(object sender, RoutedEventArgs e)
        {
            string artNr = " " + ArticleNum.Text;
            string productInfo = " " + ProductInfo.Text;
            string codeEAN = EANcode.Text;
            
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            //set up sheet margin
            sheet1.SetMargin(MarginType.LeftMargin, 0.1);
            sheet1.SetMargin(MarginType.RightMargin, 0);
            sheet1.SetMargin(MarginType.TopMargin, 0.1);
            sheet1.SetMargin(MarginType.BottomMargin, 0);

            if (cb8.IsChecked == true)
            {
                // Create a new fonts
                IFont font1 = workbook.CreateFont();
                font1.FontHeightInPoints = 10;
                font1.FontName = "Arial";
                font1.IsBold = true;

                IFont font2 = workbook.CreateFont();
                font2.FontHeightInPoints = 8;
                font2.FontName = "Arial";

                // Set fonts into new styles
                ICellStyle fontStyle1 = workbook.CreateCellStyle();
                fontStyle1.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Bottom;
                fontStyle1.SetFont(font1);
                ICellStyle fontStyle2 = workbook.CreateCellStyle();
                fontStyle2.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                fontStyle2.SetFont(font2);

                // set column width
                for (int i = 0; i < 4; i++)
                {
                    sheet1.SetColumnWidth(i, 6700);
                }

                //add columns and rows
                int x = 0;
                for (int i = 0; i < 10; i++)
                {
                    IRow row1 = sheet1.CreateRow(x);
                    row1.Height = 280;
                    x++;

                    for (int j = 0; j < 4; j++)
                    {
                        ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(artNr);
                        cell.CellStyle = fontStyle1;
                    }

                    IRow row2 = sheet1.CreateRow(x);
                    row2.Height = 250;
                    x++;

                    for (int j = 0; j < 4; j++)
                    {
                        ICell cell = row2.CreateCell(j);
                        cell.SetCellValue(productInfo);
                        cell.CellStyle = fontStyle2;
                    }

                    IRow row3 = sheet1.CreateRow(x);
                    row3.Height = 1130;
                    x++;

                    for (int j = 0; j < 4; j++)
                    {
                        ICell cell = row3.CreateCell(j);
                    }
                }

                //Generate barcode           
                MemoryStream ms1 = BarCodeManager.GenerateBarcode(codeEAN, 90, 225);

                //add barcode to .xlsx
                BarCodeManager.InsertBarcodeToSheet(31, 4, workbook, sheet1, ms1);

                //the following three statements are required only for HSSF
                //sheet1.FitToPage = (true);
                IPrintSetup printSetup = sheet1.PrintSetup;
                printSetup.PaperSize = ((short)PaperSize.A4_Small);
                //printSetup.FitHeight = ((short)1);
                //printSetup.FitWidth = ((short)1);

                //save file
                SaveManager.SaveSheet(artNr, workbook);
            }
            else if (cb40.IsChecked == true)
            {
                //set up sheet margin
                sheet1.SetMargin(MarginType.LeftMargin, 0.1);
                sheet1.SetMargin(MarginType.RightMargin, 0);
                sheet1.SetMargin(MarginType.TopMargin, 0.1);
                sheet1.SetMargin(MarginType.BottomMargin, 0);

                // Create a new fonts
                IFont font1 = workbook.CreateFont();
                font1.FontHeightInPoints = 24;
                font1.FontName = "Arial";
                font1.IsBold = true;

                IFont font2 = workbook.CreateFont();
                font2.FontHeightInPoints = 18;
                font2.FontName = "Arial";

                // Set fonts into new styles
                ICellStyle fontStyle1 = workbook.CreateCellStyle();
                fontStyle1.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Bottom;
                fontStyle1.SetFont(font1);
                ICellStyle fontStyle2 = workbook.CreateCellStyle();
                fontStyle2.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                fontStyle2.SetFont(font2);

                // set column width
                for (int i = 0; i < 2; i++) //number of columns
                {
                    sheet1.SetColumnWidth(i, 13350);
                }

                //add columns and rows
                int x = 0;
                for (int i = 0; i < 4; i++) //number of rows
                {
                    IRow row1 = sheet1.CreateRow(x);
                    row1.Height = 850;
                    x++;

                    for (int j = 0; j < 2; j++) //number of columns
                    {
                        ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(artNr);
                        cell.CellStyle = fontStyle1;
                    }

                    IRow row2 = sheet1.CreateRow(x);
                    row2.Height = 500;
                    x++;

                    for (int j = 0; j < 2; j++) //number of columns
                    {
                        ICell cell = row2.CreateCell(j);
                        cell.SetCellValue(productInfo);
                        cell.CellStyle = fontStyle2;
                    }

                    IRow row3 = sheet1.CreateRow(x);
                    row3.Height = 2800;
                    x++;

                    for (int j = 0; j < 2; j++) //number of columns
                    {
                        ICell cell = row3.CreateCell(j);
                    }
                }

                //Generate barcode           
                MemoryStream ms1 = BarCodeManager.GenerateBarcode(codeEAN, 120, 240);

                //add barcode to .xlsx
                BarCodeManager.InsertBarcodeToSheet(13, 2, 1, 0.7, workbook, sheet1, ms1);

                if (ms1 != null)
                {
                    //the following three statements are required only for HSSF
                    IPrintSetup printSetup = sheet1.PrintSetup;
                    printSetup.PaperSize = ((short)PaperSize.A4_Small);

                    //save file
                    SaveManager.SaveSheet(artNr, workbook);
                }
            }
            else if (cb5x5.IsChecked == true)
            {
                MessageBox.Show("Zatím nefunguje.");
            }
            else
            {
                MessageBox.Show("Nejdříve zvolte formát.");
            }
        }

        private void Checked5x5(object sender, RoutedEventArgs e) { cb8.IsChecked = false; cb40.IsChecked = false; }
        private void Checked8(object sender, RoutedEventArgs e) { cb5x5.IsChecked = false; cb40.IsChecked = false; }
        private void Checked40(object sender, RoutedEventArgs e) {cb8.IsChecked = false; cb5x5.IsChecked = false; }
    }
}
