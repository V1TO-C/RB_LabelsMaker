﻿using Microsoft.Win32;
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
using ZXing;
using ZXing.Common;

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

        private void Button_Click_40(object sender, RoutedEventArgs e)
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

            //add columns and rows
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
                row3.Height = 1020;
                x++;

                for (int j = 0; j < 4; j++)
                {
                    ICell cell = row3.CreateCell(j);
                }
            }

            //Generate barcode           
            BarcodeWriter writer = new BarcodeWriter()
            {
                Format = BarcodeFormat.CODE_128,
                Options = new EncodingOptions
                {
                    Height = 85,
                    Width = 230,
                    PureBarcode = false,
                    Margin = 0,
                }
            };

            //"C:/Development/RB_LabelsMaker/Sources/EAN-13.png"
            var bitmap = writer.Write("987654321234");
            MemoryStream ms = new MemoryStream();
            bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);

            //add barcode to .xlsx
            byte[] data = ms.ToArray();
            int pictureIndex = workbook.AddPicture(data, PictureType.JPEG);
            ICreationHelper helper = workbook.GetCreationHelper();

            for (int r = 2; r < 31; r += 3)
            {
                for (int c = 0; c < 4; c++)
                {
                    IDrawing drawing = sheet1.CreateDrawingPatriarch();
                    IClientAnchor anchor = helper.CreateClientAnchor();
                    anchor.Col1 = c;
                    anchor.Row1 = r;
                    IPicture picture = drawing.CreatePicture(anchor, pictureIndex);
                    picture.Resize();
                }
            }


            //the following three statements are required only for HSSF
            //sheet1.FitToPage = (true);
            //IPrintSetup printSetup = sheet1.PrintSetup;
            //printSetup.FitHeight = ((short)1);
            //printSetup.FitWidth = ((short)1);


            SaveManager.SaveSheet(workbook);
        }

        private void Button_Click_8(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("8");
        }

        private void Button_Click_5x5(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("5x5");
        }
    }
}
