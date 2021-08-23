using NPOI.SS.UserModel;
using System;
using System.IO;
using System.Windows;

namespace RB_LabelsMaker
{
    class SaveManager
    {
        public static void SaveSheet(string artNr, IWorkbook file)
        {
            //configure save file dialog box
            Microsoft.Win32.SaveFileDialog dlg = new()
            {
                FileName = artNr, //default file name
                DefaultExt = ".xlsx", //default file extension
                Filter = "XLSX documents (.xlsx)|*.xlsx" //filter files by extension
            };

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                try
                {
                    // Save document
                    string filename = dlg.FileName;
                    FileStream out1 = new(filename, FileMode.Create);
                    file.Write(out1);
                    out1.Close();
                    MessageBox.Show($"Soubor {dlg.FileName} byl uložen.");
                }
                catch (Exception)
                {
                    MessageBox.Show("Něco se pokazilo při ukládání, zkontroluj, jestli je dobře zadaný název nebo jestli není otevřený soubor se stejným názvem.");
                }
            }
        }
    }
}
