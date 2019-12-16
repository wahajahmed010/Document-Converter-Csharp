using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace Document_converter
{
    class PdfToTxt : Conversion
    {
        public void Converter(string path)
        {
            string pdfFile = path; 
            string textFile = Path.ChangeExtension(pdfFile, ".txt");

            //Convert PDF file to Text file
            SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
            //this property is necessary only for registered version
            //f.Serial = "XXXXXXXXXXX";

            f.OpenPdf(pdfFile);

            if (f.PageCount > 0)
            {
                int result = f.ToText(textFile);

                //Show Text document
                if (result == 0)
                {
                   // System.Diagnostics.Process.Start(textFile);
                    MessageBox.Show("Conversion is Completed!");
                    
                }
            }
        }
    }
}
