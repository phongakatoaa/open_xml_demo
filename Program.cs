using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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

            var processor = new WordProcessor(fileName);
            processor.FindAndReplace("#NAME#", "Phong Ha Tuan");
            processor.FindAndReplace("#DATE#", DateTime.Now.ToString());
            processor.InsertTable("TABLE_PLACEHOLDER");
            processor.InsertImage("IMAGE_PLACEHOLDER", projectDirectory + "\\demo\\image.png", 200, 200);
            processor.FindAndReplace("#USERNAME#", "Hoang Anh Tu", new Color() { Val = "FF0000" });
            processor.FindAndReplace("#TASK#", "Do something", new Color() { Val = "0AC22D" });
            processor.Save(saveLoc);
        }
    }
}
