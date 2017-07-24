using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CSSPFCFormWriterDLL.Services
{
    public partial class CSSPFCFormWriter
    {
        public void DoParagraphBelowTable1(Paragraph paragraph7)
        {
            ParagraphProperties paragraphProperties7 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts15 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize2 = new FontSize() { Val = "8" };

            paragraphMarkRunProperties7.Append(runFonts15);
            paragraphMarkRunProperties7.Append(fontSize2);

            paragraphProperties7.Append(paragraphMarkRunProperties7);

            paragraph7.Append(paragraphProperties7);
        }
    }
}
