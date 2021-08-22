using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
                Format = BarcodeFormat.CODE_128,
                Options = new EncodingOptions
                {
                    Height = codeHeight,
                    Width = codeWidth,
                    PureBarcode = false,
                    Margin = 0,
                }
            };
            var bitmap = writer.Write(code);
            MemoryStream ms = new MemoryStream();
            bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
            return ms;
        }

        public static void InsertBarcodeToSheet(int rowNum, int colNum, double scale, IWorkbook wb, ISheet sheet, MemoryStream ms)
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
                    IPicture picture = drawing.CreatePicture(anchor, pictureIndex);
                    picture.Resize(scale);
                }
            }
        }
    }
}
