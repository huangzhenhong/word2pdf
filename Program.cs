using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.IO;
using System.Text;

namespace word2pdf
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            var basePath = @"C:\Users\dannier_huang\Desktop\tests\";
            //var fileInfo = new FileInfo(basePath + @"ExampleOutput-20-07-07-170637\BrochureTemplate-3-ConvertedByHtmlToWml.docx");
            //string fullFilePath = fileInfo.FullName;
            //string htmlText = string.Empty;
            //try
            //{
            //    htmlText = HTMLService.ParseDOCX(fileInfo);
            //}
            //catch (OpenXmlPackageException e)
            //{
            //    if (e.ToString().Contains("Invalid Hyperlink"))
            //    {
            //        using (FileStream fs = new FileStream(fullFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            //        {
            //            UriFixer.FixInvalidUri(fs, brokenUri => HTMLService.FixUri(brokenUri));
            //        }
            //        htmlText = HTMLService.ParseDOCX(fileInfo);
            //    }
            //}

            //var writer = File.CreateText(basePath + @"converted\test1.html");
            //writer.WriteLine(htmlText.ToString());
            //writer.Dispose();

            //Console.WriteLine("html file generated");

            //var builder = new StringBuilder();
            //using (StreamReader reader = System.IO.File.OpenText(basePath + "test1.html"))
            //{
            //    builder.Append(reader.ReadToEnd());
            //}

            //PDFService.GeneratePDF(builder.ToString(), basePath + @"converted\ECO_Template.pdf");

            //Console.WriteLine("Pdf generated");


            // Convert html to word 

            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(basePath + string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            foreach (var file in Directory.GetFiles(@"C:\Users\dannier_huang\Desktop\tests\", "*.html"))
            {
                HTMLService.ConvertToDocx(file, tempDi.FullName);
            }

            //HTMLService.ConvertToDocx(basePath + "BrochureTemplate.html", basePath + @"converted\BrochureTemplate.docx");


        }
    }
}
