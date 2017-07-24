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
        public void DoParagraph1(Paragraph paragraph1)
        {
            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold1 = new Bold();

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(bold1);

            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run() { RsidRunProperties = "0005391D" };

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold2 = new Bold();

            runProperties1.Append(runFonts2);
            runProperties1.Append(bold2);
            Text text1 = new Text();
            text1.Text = "A1 Fecal Coliform - Water Analysis Data Sheet";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
        }
    }
}
