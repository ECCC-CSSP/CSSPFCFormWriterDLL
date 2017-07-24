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
        public void DoParagraphBelowTable2(Paragraph paragraph69)
        {
            ParagraphProperties paragraphProperties69 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties69 = new ParagraphMarkRunProperties();
            Vanish vanish1 = new Vanish();
            FontSize fontSize118 = new FontSize() { Val = "20" };

            paragraphMarkRunProperties69.Append(vanish1);
            paragraphMarkRunProperties69.Append(fontSize118);

            paragraphProperties69.Append(paragraphMarkRunProperties69);

            Run run63 = new Run() { RsidRunProperties = "00D10A17" };

            RunProperties runProperties63 = new RunProperties();
            RunFonts runFonts131 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize119 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript117 = new FontSizeComplexScript() { Val = "22" };

            runProperties63.Append(runFonts131);
            runProperties63.Append(fontSize119);
            runProperties63.Append(fontSizeComplexScript117);
            Text text63 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text63.Text = "RECORDED TEMPERATURE IS ACTUAL READING + CORRECTION ";

            run63.Append(runProperties63);
            run63.Append(text63);
            ProofError proofError45 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run64 = new Run() { RsidRunProperties = "00D10A17" };

            RunProperties runProperties64 = new RunProperties();
            RunFonts runFonts132 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize120 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript118 = new FontSizeComplexScript() { Val = "22" };

            runProperties64.Append(runFonts132);
            runProperties64.Append(fontSize120);
            runProperties64.Append(fontSizeComplexScript118);
            Text text64 = new Text();
            text64.Text = "FACTOR";

            run64.Append(runProperties64);
            run64.Append(text64);

            paragraph69.Append(paragraphProperties69);
            paragraph69.Append(run63);
            paragraph69.Append(proofError45);
            paragraph69.Append(run64);
        }
    }
}
