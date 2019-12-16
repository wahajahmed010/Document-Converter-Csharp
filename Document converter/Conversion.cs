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
    class Conversion
    {
        public void fileconvert(string filepath, string savepath, string openwali, string savewali, string filname, string sfilename)
        {

            //Checking the files that are to be converted.

            
            //int a = openFileDialog3.FilterIndex;
            //int b = saveFileDialog3.FilterIndex;

            // MessageBox.Show("Error: Same extension selected");

            if (openwali == ".doc") //.doc
            {
                if (savewali == ".docx") //.docx
                {
                    MessageBox.Show("Error: Same file extension");
                }
                else if (savewali == ".doc") //.doc
                {
                    MessageBox.Show("Error: Same file extension");
                }
                else if (savewali == ".pdf") //.pdf
                {

                    DocToPdf d2p = new DocToPdf();
                    d2p.Convertor(filepath, savepath);

                }
                else if (savewali == ".txt") //.txt
                {
                    DocToTxt DtT = new DocToTxt();
                    DtT.Converter(filepath, savepath);
                    //MessageBox.Show("Abhi rukja ye nhi hoa");
                }



            }
            else if (openwali == ".docx") //.docx
            {
                if (savewali == ".docx") //.docx
                {
                    MessageBox.Show("Error: Same file extension");
                }
                else if (savewali == ".doc") //.doc
                {
                    MessageBox.Show("Error: Same file extension");
                }
                else if (savewali == ".pdf") //.pdf
                {

                    DocToPdf d2p = new DocToPdf();
                    d2p.Convertor(filepath, savepath);

                }
                else if (savewali == ".txt") //.txt
                {
                    DocToTxt DtT = new DocToTxt();
                    DtT.Converter(filepath, savepath);
                    //MessageBox.Show("Abhi rukja bhai");
                }

            }
            else if (openwali == ".pdf") //.pdf
            {
                if (savewali == ".docx") //.docx
                {
                    PdfToWord ptw = new PdfToWord();
                    ptw.Converter(filepath, savepath);

                }
                else if (savewali == ".doc") //.doc
                {
                    PdfToWord ptw = new PdfToWord();
                    ptw.Converter(filepath, savepath);
                }
                else if (savewali == ".pdf") //.pdf
                {
                    MessageBox.Show("Error: Same file extension");

                }
                else if (savewali == ".txt") //.txt
                {
                    PdfToTxt PtT = new PdfToTxt();
                    PtT.Converter(filepath);
                }

            }
            
                


        }
    }
}
