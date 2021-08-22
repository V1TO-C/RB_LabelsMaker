using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace RB_LabelsMaker
{
    class SaveManager
    {
        public static void SaveSheet(IWorkbook file)
        {
            //configure save file dialog box
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "Document"; //default file name
            dlg.DefaultExt = ".xlsx"; //default file extension
            dlg.Filter = "XLSX documents (.xlsx)|*.xlsx"; //filter files by extension

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                string filename = dlg.FileName;
                FileStream out1 = new FileStream(filename, FileMode.Create);
                file.Write(out1);
                out1.Close();

                MessageBox.Show($"Soubor {dlg.FileName} byl uložen.");
            }
        }
    }
}
