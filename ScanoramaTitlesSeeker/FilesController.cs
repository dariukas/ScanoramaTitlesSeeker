using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScanoramaTitlesSeeker
{
    class FilesController
    {
        //default folder path
        public static String folderPath = "";

        public static void createDeaultFolder(string defaultFolderTitle) {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            folderPath = desktopPath + "\\"+ defaultFolderTitle;
            bool exists = System.IO.Directory.Exists(folderPath);
            if (!exists) { 
            System.IO.Directory.CreateDirectory(folderPath);
                System.Console.WriteLine("The default directory was created in {0}", folderPath);
            } else {
                System.Console.WriteLine("The default directory already exists at {0}", folderPath);
            }
        }

        //open the titles file using the dialog menu
        public static string openFile()
        {
            string filePath = "";
            //Browse Files
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.InitialDirectory = folderPath;
            dlg.DefaultExt = ".pptx";
            dlg.Filter = "PPTX Files (*.pptx)|*.pptx|PPT Files (*.ppt)|*.ppt";
            Nullable<bool> result = dlg.ShowDialog();
            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                filePath = dlg.FileName;
                //textBox1.Text = filename;
            }
            else
            {
                System.Console.WriteLine("Couldn't show the dialog.");
            }
            return filePath;
        }

        public static string openTxtFile()
        {
            //string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = "";
            //Browse Files
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.InitialDirectory = folderPath;
            dlg.DefaultExt = ".txt";
            dlg.Filter = "TXT Files (*.txt)|*.txt";
            Nullable<bool> result = dlg.ShowDialog();
            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                filePath = dlg.FileName;
            }
            else
            {
                System.Console.WriteLine("Couldn't show the dialog.");
            }
            return filePath;
        }

        //save to .txt file using the dialog
        public static string saveTxtFile(String fileName)
        {
            string filePath = "";
            //Browse Files
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.InitialDirectory = folderPath;
            dlg.DefaultExt = ".txt";
            dlg.Filter = "TXT Files (*.txt)|*.txt";
            dlg.FileName = fileName;
            Nullable<bool> result = dlg.ShowDialog();
            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                filePath = dlg.FileName;
            }
            else
            {
                System.Console.WriteLine("Couldn't show the dialog.");
            }
            System.Console.WriteLine("Titles will be saved in {0}.", filePath);
            return filePath;
        }

        public static void saveSlidesInBackground(Microsoft.Office.Interop.PowerPoint.Presentation presentation, String fileName) {
            String filePath = folderPath+fileName+".pptx";
            presentation.SaveAs(filePath,
            Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
            Microsoft.Office.Core.MsoTriState.msoTriStateMixed);
            System.Console.WriteLine("PowerPoint application saved in {0}.", filePath);
        }

        public static void saveSlides(Microsoft.Office.Interop.PowerPoint.Presentation presentation)
        {
            string filePath ="";
            // Microsoft.Office.Interop.PowerPoint.FileConverter fc = new Microsoft.Office.Interop.PowerPoint.FileConverter();
            // if (fc.CanSave) { }
            //https://msdn.microsoft.com/en-us/library/system.windows.forms.savefiledialog(v=vs.110).aspx

            //Browse Files
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.InitialDirectory = folderPath;
            dlg.DefaultExt = ".pptx";
            dlg.Filter = "PPTX Files (*.pptx)|*.pptx";
            Nullable<bool> result = dlg.ShowDialog();
            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document
                filePath = dlg.FileName;
                //textBox1.Text = filename;
            }
            else
            {
                System.Console.WriteLine("Couldn't show the dialog.");
            }
            presentation.SaveAs(filePath,
            Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
            Microsoft.Office.Core.MsoTriState.msoTriStateMixed);
            System.Console.WriteLine("PowerPoint application was saved in {0}.", filePath);
        }
    }
}
