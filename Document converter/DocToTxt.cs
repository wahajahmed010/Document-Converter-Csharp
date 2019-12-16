using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using Word = Microsoft.Office.Interop.Word;


namespace Document_converter
{
    class DocToTxt : Conversion
    {
        //Microsoft.Office.Interop.Word.Document wordDoc { get; set; }
        public void Converter(string docPath, string savePath)
        {
            
            object path = docPath;
            string txtPath = savePath;
            Word.Application app = new Word.Application();
            Word.Document doc;
            object missing = Type.Missing;
            object readOnly = true;
            try
            {
                doc = app.Documents.Open(ref path, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                string text = doc.Content.Text;
                File.WriteAllText(txtPath, text);
               
            }
            catch
            {
                MessageBox.Show("An error occured. Please check the file path to your word document, and whether the word document is valid.");
            }
            finally
            {
                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                app.Quit(ref saveChanges, ref missing, ref missing);
            }


            MessageBox.Show("Conversion is Completed!");
            
            
            
            
            
            
            
       


            


        }
    }
}
