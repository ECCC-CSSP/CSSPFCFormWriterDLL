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
        public void DoParagraphBelowTable3_2(Paragraph paragraph474)
        {
            ParagraphProperties paragraphProperties474 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties474 = new ParagraphMarkRunProperties();
            RunFonts runFonts603 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize262 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript260 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties474.Append(runFonts603);
            paragraphMarkRunProperties474.Append(fontSize262);
            paragraphMarkRunProperties474.Append(fontSizeComplexScript260);

            paragraphProperties474.Append(paragraphMarkRunProperties474);

            Run run131 = new Run() { RsidRunProperties = "00D10A17" };

            RunProperties runProperties131 = new RunProperties();
            RunFonts runFonts604 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize263 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript261 = new FontSizeComplexScript() { Val = "22" };

            runProperties131.Append(runFonts604);
            runProperties131.Append(fontSize263);
            runProperties131.Append(fontSizeComplexScript261);
            Text text131 = new Text();
            text131.Text = "Note: All required information as per section 5.10.2 of ISO/IEC 17025 is available from the Laboratory Supervisor.";

            run131.Append(runProperties131);
            run131.Append(text131);

            paragraph474.Append(paragraphProperties474);
            paragraph474.Append(run131);
        }
    }
}
