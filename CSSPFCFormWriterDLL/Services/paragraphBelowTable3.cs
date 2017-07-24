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
        public void DoParagraphBelowTable3(Paragraph paragraph473)
        {
            ParagraphProperties paragraphProperties473 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties473 = new ParagraphMarkRunProperties();
            RunFonts runFonts602 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize261 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript259 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties473.Append(runFonts602);
            paragraphMarkRunProperties473.Append(fontSize261);
            paragraphMarkRunProperties473.Append(fontSizeComplexScript259);

            paragraphProperties473.Append(paragraphMarkRunProperties473);

            paragraph473.Append(paragraphProperties473);

            Paragraph paragraph474 = new Paragraph() { RsidParagraphMarkRevision = "00D10A17", RsidParagraphAddition = "00004340", RsidRunAdditionDefault = "00D53280" };

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
