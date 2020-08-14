using DinkToPdf;
using DinkToPdf.Contracts;
using System.IO;
using System.Text;

namespace word2pdf
{
    public static class PDFService
    {
        public static void GeneratePDF(string html, string outPath)
        {
            var doc = new HtmlToPdfDocument()
            {
                GlobalSettings = {
                    ColorMode = ColorMode.Color,
                    Orientation = Orientation.Portrait,
                    PaperSize = PaperKind.A4Plus,
                    Margins = new MarginSettings() { Top = 10 },
                    Out = outPath,
                },
                Objects = {
                new ObjectSettings() {
                    PagesCount = true,
                    HtmlContent = html,
                    WebSettings = { DefaultEncoding = "utf-8" },
                    HeaderSettings = { FontSize = 9, Right = "Page [page] of [toPage]", Line = true, Spacing = 2.812 }
                    }
                }
            };

            var converter = new BasicConverter(new PdfTools());
            converter.Convert(doc);
        }

    }
}
