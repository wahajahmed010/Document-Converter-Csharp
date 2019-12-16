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

namespace Document_converter
{
    class DocToPdf : Conversion
    {
        Microsoft.Office.Interop.Word.Document wordDoc { get; set; }

        public void Convertor(string loadpath, string savepath)
        {
            
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            wordDoc = app.Documents.Open(loadpath);
            wordDoc.ExportAsFixedFormat(savepath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);

            wordDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            app.Quit();

            
        
            Marshal.ReleaseComObject(wordDoc);
            Marshal.ReleaseComObject(app);

           
            MessageBox.Show("Conversion is Completed!");


        }

    }
}
