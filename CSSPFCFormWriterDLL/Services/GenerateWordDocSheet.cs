using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using M = DocumentFormat.OpenXml.Math;
using A = DocumentFormat.OpenXml.Drawing;
//using System.Windows.Forms;
//using CSSPWQInputTool;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;

namespace CSSPFCFormWriterDLL.Services
{
    public partial class CSSPFCFormWriter
    {
        private void CreateWaterAnalysisDataSheet(Body body, int SheetNumber)
        {
            List<string> AllowableTideTextList = new List<string>()
            {
                "LT", "LR", "LF", "MT", "MR", "MF", "HT", "HR", "HF"
            };

            Document document1 = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            document1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Body body1 = new Body();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00B67DA8", RsidParagraphAddition = "009A75C9", RsidParagraphProperties = "009A75C9", RsidRunAdditionDefault = "006D4BC2" };
            DoParagraph1(paragraph1);

            Table table1 = new Table();
            DoTable1(table1);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "00D10A17", RsidParagraphAddition = "00004340", RsidRunAdditionDefault = "00004340" };
            DoParagraphBelowTable1(paragraph7);

            Table table2 = new Table();
            DoTable2(table2);

            Paragraph paragraph69 = new Paragraph() { RsidParagraphMarkRevision = "00D10A17", RsidParagraphAddition = "00ED7FB2", RsidParagraphProperties = "00ED7FB2", RsidRunAdditionDefault = "00D53280" };
            DoParagraphBelowTable2(paragraph69);

            Table table3 = new Table();
            DoTable3(table3, SheetNumber);

            Paragraph paragraph473 = new Paragraph() { RsidParagraphAddition = "0032114E", RsidRunAdditionDefault = "0032114E" };
            DoParagraphBelowTable3(paragraph473);

            Paragraph paragraph474 = new Paragraph() { RsidParagraphMarkRevision = "00D10A17", RsidParagraphAddition = "00004340", RsidRunAdditionDefault = "00D53280" };
            DoParagraphBelowTable3_2(paragraph474);

            Table table4 = new Table();
            DoTable4(table4);

            Paragraph paragraph495 = new Paragraph() { RsidParagraphMarkRevision = "00584876", RsidParagraphAddition = "00584876", RsidParagraphProperties = "002164F2", RsidRunAdditionDefault = "00584876" };

            SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "00584876", RsidR = "00584876", RsidSect = "00D53280" };
            HeaderReference headerReference1 = new HeaderReference() { Type = HeaderFooterValues.Default, Id = "rId7" };
            FooterReference footerReference1 = new FooterReference() { Type = HeaderFooterValues.Default, Id = "rId8" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)15842U, Height = (UInt32Value)12242U, Orient = PageOrientationValues.Landscape, Code = (UInt16Value)1U };
            PageMargin pageMargin1 = new PageMargin() { Top = 112, Right = (UInt32Value)170U, Bottom = 284, Left = (UInt32Value)567U, Header = (UInt32Value)284U, Footer = (UInt32Value)431U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "720" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(headerReference1);
            sectionProperties1.Append(footerReference1);
            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(paragraph1);
            body1.Append(table1);
            body1.Append(paragraph7);
            body1.Append(table2);
            body1.Append(paragraph69);
            body1.Append(table3);
            body1.Append(paragraph473);
            body1.Append(paragraph474);
            body1.Append(table4);
            body1.Append(paragraph495);
            body1.Append(sectionProperties1);

            body.Append(body1);
        }
    }
}
