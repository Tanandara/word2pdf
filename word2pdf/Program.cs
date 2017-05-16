using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace word2pdf
{
    /**
     * ดัดแปลงจาก
     * https://www.codeproject.com/Questions/346784/How-to-convert-word-document-to-pdf-in-Csharp
     * 
     */
    class Program
    {
        static public Document wordDocument { get; set; }
        static void Main(string[] args)
        {
            string filepath = Directory.GetCurrentDirectory(); // ดึง path ปัจจุบัน
            DirectoryInfo d = new DirectoryInfo(filepath); // ดึงค่าจาก folder

            // ลูป docx -> pdf
            foreach (var file in d.GetFiles("*.docx"))
            {
                Console.WriteLine("Converting : " + file.Name);
                Application appWord = new Application();

                // เปิดไฟล์ word
                wordDocument = appWord.Documents.Open(filepath + "\\" + file.Name); 

                // convert ไฟล์
                wordDocument.ExportAsFixedFormat(filepath + "\\" + Path.GetFileNameWithoutExtension(file.Name) + ".pdf", WdExportFormat.wdExportFormatPDF); 
            }
            Console.WriteLine("======= Finish ========");
            Console.ReadKey();
        }

    }
}



