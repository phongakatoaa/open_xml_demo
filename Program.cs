using DocumentFormat.OpenXml.Packaging;
using open_xml_demo.util;
using System;
using System.IO;

namespace open_xml_demo
{
    class Program
    {
        static void Main(string[] args)
        {
            string workingDirectory = Environment.CurrentDirectory;
            string projectDirectory = Directory.GetParent(workingDirectory).Parent.Parent.FullName;

            string fileName = projectDirectory + "\\demo\\test.docx";
            string saveLoc = projectDirectory + "\\demo\\result.docx";

            byte[] byteArray = File.ReadAllBytes(fileName);
            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);

                var processor = new WordProcessor(stream);
                processor.FindAndReplace("#NAME#", "Phong Ha Tuan");
                processor.FindAndReplace("#DATE#", DateTime.Now.ToString());
                processor.InsertTable("TABLE_PLACEHOLDER");
                processor.Save(saveLoc);
            }
        }
    }
}
