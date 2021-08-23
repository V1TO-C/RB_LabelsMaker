using NPOI.SS.UserModel;
using NPOI.Util;
using System.IO;
using System.Windows;
using ZXing;
using ZXing.Common;


namespace RB_LabelsMaker
{
    class BarCodeManager
    {
        public static MemoryStream GenerateBarcode(string code, int codeHeight, int codeWidth)
        {
            BarcodeWriter writer = new BarcodeWriter()
            {
                Format = BarcodeFormat.EAN_13,
                Options = new EncodingOptions
                {
                    Height = codeHeight,
                    Width = codeWidth,
                    PureBarcode = false,
                    Margin = 0,
                }
            };
            try
            {
                var bitmap = writer.Write(code);
                MemoryStream ms = new MemoryStream();
                bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                return ms;
            }
            catch (System.Exception)
            {

                MessageBox.Show("Zkontroluj EAN kód, není správně vyplněný.");
                return null;
            }
        }

        public static void InsertBarcodeToSheet(int rowNum, int colNum, IWorkbook wb, ISheet sheet, MemoryStream ms)
        {
            if (ms == null)
            {

            }
            else
            {
                //add barcode to .xlsx
                byte[] data = ms.ToArray();
                int pictureIndex = wb.AddPicture(data, PictureType.JPEG);
                ICreationHelper helper = wb.GetCreationHelper();

                for (int r = 2; r < rowNum; r += 3)
                {
                    for (int c = 0; c < colNum; c++)
                    {
                        IDrawing drawing = sheet.CreateDrawingPatriarch();
                        IClientAnchor anchor = helper.CreateClientAnchor();
                        anchor.Col1 = c;
                        anchor.Row1 = r;
                        anchor.Dx1 = (Units.ToEMU(2));
                        IPicture picture = drawing.CreatePicture(anchor, pictureIndex);
                        picture.Resize();
                    }
                }
            }
        }

        public static void InsertBarcodeToSheet(int rowNum, int colNum, double resizeX, double resizeY, IWorkbook wb, ISheet sheet, MemoryStream ms)
        {
            if (ms == null)
            {
                
            }
            else
            {
                //add barcode to .xlsx
                byte[] data = ms.ToArray();
                int pictureIndex = wb.AddPicture(data, PictureType.JPEG);
                ICreationHelper helper = wb.GetCreationHelper();

                for (int r = 2; r < rowNum; r += 3)
                {
                    for (int c = 0; c < colNum; c++)
                    {
                        IDrawing drawing = sheet.CreateDrawingPatriarch();
                        IClientAnchor anchor = helper.CreateClientAnchor();
                        anchor.Col1 = c;
                        anchor.Row1 = r;
                        anchor.Dx1 = (Units.ToEMU(10));
                        IPicture picture = drawing.CreatePicture(anchor, pictureIndex);
                        picture.Resize(resizeX, resizeY);
                    }
                }
            }
        }
    }
}
