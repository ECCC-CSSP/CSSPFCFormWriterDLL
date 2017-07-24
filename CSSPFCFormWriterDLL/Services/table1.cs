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
        public void DoTable1(Table table1)
        {
            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "15120", Type = TableWidthUnitValues.Dxa };
            TableIndentation tableIndentation1 = new TableIndentation() { Width = -72, Type = TableWidthUnitValues.Dxa };
            TableLook tableLook1 = new TableLook() { Val = "01E0" };

            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableIndentation1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "1173" };
            GridColumn gridColumn2 = new GridColumn() { Width = "5103" };
            GridColumn gridColumn3 = new GridColumn() { Width = "283" };
            GridColumn gridColumn4 = new GridColumn() { Width = "3402" };
            GridColumn gridColumn5 = new GridColumn() { Width = "5159" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);
            tableGrid1.Append(gridColumn5);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "006B4849", RsidTableRowProperties = "00FA3A80" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "1173", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(topBorder1);
            tableCellBorders1.Append(leftBorder1);
            tableCellBorders1.Append(bottomBorder1);
            tableCellBorders1.Append(rightBorder1);
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(tableCellBorders1);
            tableCellProperties1.Append(shading1);
            tableCellProperties1.Append(tableCellVerticalAlignment1);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "006B4849", RsidParagraphProperties = "006B4849", RsidRunAdditionDefault = "006B4849" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties2.Append(runFonts3);

            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run();

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            runProperties2.Append(runFonts4);
            Text text2 = new Text();
            text2.Text = "Location";

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph2);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "5103", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(topBorder2);
            tableCellBorders2.Append(leftBorder2);
            tableCellBorders2.Append(bottomBorder2);
            tableCellBorders2.Append(rightBorder2);
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(tableCellBorders2);
            tableCellProperties2.Append(shading2);
            tableCellProperties2.Append(tableCellVerticalAlignment2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "006B4849", RsidParagraphProperties = "00030688", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties3.Append(runFonts5);

            paragraphProperties3.Append(justification2);
            paragraphProperties3.Append(paragraphMarkRunProperties3);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run3 = new Run();

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            runProperties3.Append(runFonts6);
            Text text3 = new Text();
            string Location = labSheetA1Sheet.SubsectorName.Substring(labSheetA1Sheet.SubsectorName.IndexOf(" ") + 1);
            text3.Text = labSheetA1Sheet.SubsectorName.Substring(0, labSheetA1Sheet.SubsectorName.IndexOf(" ")) +
                " (" + (Location.Length > 20 ? Location.Substring(0, 20).Replace("(", "").Replace(")", "") + " ...)" : Location.Replace("(", "").Replace(")", "") + ")");

            run3.Append(runProperties3);
            run3.Append(text3);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(proofError1);
            paragraph3.Append(run3);
            paragraph3.Append(proofError2);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph3);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "283", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(leftBorder3);
            tableCellBorders3.Append(rightBorder3);
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellBorders3);
            tableCellProperties3.Append(shading3);
            tableCellProperties3.Append(tableCellVerticalAlignment3);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "008162D3", RsidParagraphAddition = "006B4849", RsidParagraphProperties = "008162D3", RsidRunAdditionDefault = "006B4849" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize1 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties4.Append(runFonts7);
            paragraphMarkRunProperties4.Append(fontSize1);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript1);

            paragraphProperties4.Append(paragraphMarkRunProperties4);

            paragraph4.Append(paragraphProperties4);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph4);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "3402", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders4.Append(topBorder3);
            tableCellBorders4.Append(leftBorder4);
            tableCellBorders4.Append(bottomBorder3);
            tableCellBorders4.Append(rightBorder4);
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(tableCellBorders4);
            tableCellProperties4.Append(shading4);
            tableCellProperties4.Append(tableCellVerticalAlignment4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "006B4849", RsidParagraphProperties = "008162D3", RsidRunAdditionDefault = "006B4849" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties5.Append(runFonts8);

            paragraphProperties5.Append(justification3);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run4 = new Run() { RsidRunProperties = "006B4849" };

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Highlight highlight1 = new Highlight() { Val = HighlightColorValues.Yellow };

            runProperties4.Append(runFonts9);
            runProperties4.Append(highlight1);
            Text text4 = new Text();
            text4.Text = "Sampling / Processing Date:";

            run4.Append(runProperties4);
            run4.Append(text4);

            Run run5 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            runProperties5.Append(runFonts10);
            Text text5 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text5.Text = "  ";

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run4);
            paragraph5.Append(run5);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph5);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "5159", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders5 = new TableCellBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder5 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder5 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders5.Append(topBorder4);
            tableCellBorders5.Append(leftBorder5);
            tableCellBorders5.Append(bottomBorder4);
            tableCellBorders5.Append(rightBorder5);
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(tableCellBorders5);
            tableCellProperties5.Append(shading5);
            tableCellProperties5.Append(tableCellVerticalAlignment5);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "006B4849", RsidParagraphProperties = "00030688", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color1 = new Color() { Val = "EAEAEA" };

            paragraphMarkRunProperties6.Append(runFonts11);
            paragraphMarkRunProperties6.Append(color1);

            paragraphProperties6.Append(justification4);
            paragraphProperties6.Append(paragraphMarkRunProperties6);
            ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run6 = new Run();

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            runProperties6.Append(runFonts12);
            Text text6 = new Text();
            DateTime SampleDate = new DateTime(int.Parse(labSheetA1Sheet.RunYear), int.Parse(labSheetA1Sheet.RunMonth), int.Parse(labSheetA1Sheet.RunDay));
            DateTime ProcessDate = (bool.Parse(labSheetA1Sheet.IncubationStartSameDay) == true ? SampleDate : SampleDate.AddDays(1));
            text6.Text = SampleDate.ToString("yyyy MMMM dd") + " / " + ProcessDate.ToString("yyyy MMMM dd");

            run6.Append(runProperties6);
            run6.Append(text6);
            ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.SpellEnd };


            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(proofError3);
            paragraph6.Append(run6);
            paragraph6.Append(proofError4);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph6);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(tableCell4);
            tableRow1.Append(tableCell5);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "00D10A17", RsidParagraphAddition = "00004340", RsidRunAdditionDefault = "00004340" };

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
