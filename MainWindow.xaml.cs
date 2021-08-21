using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Drawing;
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
using Color = System.Drawing.Color;

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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            // set column width
            for (int i = 0; i < 4; i++)
            {
                sheet1.SetColumnWidth(i, 6850);
            }

            sheet1.SetMargin(MarginType.LeftMargin, 0);
            sheet1.SetMargin(MarginType.RightMargin, 0);
            sheet1.SetMargin(MarginType.TopMargin, 0);
            sheet1.SetMargin(MarginType.BottomMargin, 0);
            

            // Create a new font and alter it
            IFont font1 = workbook.CreateFont();
            font1.FontHeightInPoints = 10;
            font1.FontName = "Arial";
            font1.IsBold = true;

            IFont font2 = workbook.CreateFont();
            font2.FontHeightInPoints = 8;
            font2.FontName = "Arial";


            // Fonts are set into a style so create a new one to use.
            ICellStyle fontStyle1 = workbook.CreateCellStyle();
            fontStyle1.SetFont(font1);
            ICellStyle fontStyle2 = workbook.CreateCellStyle();
            fontStyle2.SetFont(font2);


            int x = 0;
            for (int i = 0; i < 10; i++)
            {
                IRow row1 = sheet1.CreateRow(x);
                row1.Height = 300;
                x++;

                for (int j = 0; j < 4; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue("Art. Nr. 407-80");
                    cell.CellStyle = fontStyle1;
                }

                IRow row2 = sheet1.CreateRow(x);
                row2.Height = 250;
                x++;

                for (int j = 0; j < 4; j++)
                {
                    ICell cell = row2.CreateCell(j);
                    cell.SetCellValue("Stillkissen190cm Tierchentürkis");
                    cell.CellStyle = fontStyle2;
                }

                IRow row3 = sheet1.CreateRow(x);
                row3.Height = 1125;
                x++;

                for (int j = 0; j < 4; j++)
                {
                    ICell cell = row3.CreateCell(j);
                }
            }

            //Generate and add barcode to cells
            BarcodeLib.Barcode b = new BarcodeLib.Barcode();
            System.Drawing.Image img = b.Encode(BarcodeLib.TYPE.EAN13, "038000356216", Color.Black, Color.White, 120, 90);
            img.Save("C:/Users/vcerm/source/repos/RB_LabelsMaker/Sources/EAN-13.png", System.Drawing.Imaging.ImageFormat.Png);
            byte[] data = File.ReadAllBytes("C:/Users/vcerm/source/repos/RB_LabelsMaker/Sources/EAN-13.png");
            int pictureIndex = workbook.AddPicture(data, PictureType.JPEG);
            ICreationHelper helper = workbook.GetCreationHelper();
            IDrawing drawing = sheet1.CreateDrawingPatriarch();
            IClientAnchor anchor = helper.CreateClientAnchor();
            anchor.Col1 = 2;
            anchor.Row1 = 3;
            IPicture picture = drawing.CreatePicture(anchor, pictureIndex);
            picture.Resize();

            //the following three statements are required only for HSSF
            //sheet1.FitToPage = (true);
            //IPrintSetup printSetup = sheet1.PrintSetup;
            //printSetup.FitHeight = ((short)1);
            //printSetup.FitWidth = ((short)1);



            FileStream out1 = new FileStream("C:/Users/vcerm/source/repos/RB_LabelsMaker/Sources/table.xlsx", FileMode.Create);
            workbook.Write(out1);
            out1.Close();


            

            //MessageBox.Show("wtf");
        }
    }
}
