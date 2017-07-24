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
        public void DoTable2(Table table2)
        {
            TableProperties tableProperties2 = new TableProperties();
            TableWidth tableWidth2 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableIndentation tableIndentation2 = new TableIndentation() { Width = -72, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder5 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder6 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder6 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder5);
            tableBorders1.Append(leftBorder6);
            tableBorders1.Append(bottomBorder5);
            tableBorders1.Append(rightBorder6);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLayout tableLayout1 = new TableLayout() { Type = TableLayoutValues.Fixed };
            TableLook tableLook2 = new TableLook() { Val = "01E0" };

            tableProperties2.Append(tableWidth2);
            tableProperties2.Append(tableIndentation2);
            tableProperties2.Append(tableBorders1);
            tableProperties2.Append(tableLayout1);
            tableProperties2.Append(tableLook2);

            TableGrid tableGrid2 = new TableGrid();
            GridColumn gridColumn6 = new GridColumn() { Width = "1031" };
            GridColumn gridColumn7 = new GridColumn() { Width = "992" };
            GridColumn gridColumn8 = new GridColumn() { Width = "284" };
            GridColumn gridColumn9 = new GridColumn() { Width = "1417" };
            GridColumn gridColumn10 = new GridColumn() { Width = "1985" };
            GridColumn gridColumn11 = new GridColumn() { Width = "283" };
            GridColumn gridColumn12 = new GridColumn() { Width = "1416" };
            GridColumn gridColumn13 = new GridColumn() { Width = "238" };
            GridColumn gridColumn14 = new GridColumn() { Width = "1031" };
            GridColumn gridColumn15 = new GridColumn() { Width = "906" };
            GridColumn gridColumn16 = new GridColumn() { Width = "1087" };
            GridColumn gridColumn17 = new GridColumn() { Width = "1007" };
            GridColumn gridColumn18 = new GridColumn() { Width = "906" };
            GridColumn gridColumn19 = new GridColumn() { Width = "906" };
            GridColumn gridColumn20 = new GridColumn() { Width = "1631" };

            tableGrid2.Append(gridColumn6);
            tableGrid2.Append(gridColumn7);
            tableGrid2.Append(gridColumn8);
            tableGrid2.Append(gridColumn9);
            tableGrid2.Append(gridColumn10);
            tableGrid2.Append(gridColumn11);
            tableGrid2.Append(gridColumn12);
            tableGrid2.Append(gridColumn13);
            tableGrid2.Append(gridColumn14);
            tableGrid2.Append(gridColumn15);
            tableGrid2.Append(gridColumn16);
            tableGrid2.Append(gridColumn17);
            tableGrid2.Append(gridColumn18);
            tableGrid2.Append(gridColumn19);
            tableGrid2.Append(gridColumn20);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00397699", RsidTableRowProperties = "0032114E" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)230U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "1031", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge1 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders6 = new TableCellBorders();
            TopBorder topBorder6 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder7 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder7 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders6.Append(topBorder6);
            tableCellBorders6.Append(leftBorder7);
            tableCellBorders6.Append(rightBorder7);
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(verticalMerge1);
            tableCellProperties6.Append(tableCellBorders6);
            tableCellProperties6.Append(shading6);
            tableCellProperties6.Append(tableCellVerticalAlignment6);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            Justification justification5 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            RunFonts runFonts16 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize3 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties8.Append(runFonts16);
            paragraphMarkRunProperties8.Append(fontSize3);
            paragraphMarkRunProperties8.Append(fontSizeComplexScript2);

            paragraphProperties8.Append(justification5);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            Run run9 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize4 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "20" };

            runProperties9.Append(runFonts17);
            runProperties9.Append(fontSize4);
            runProperties9.Append(fontSizeComplexScript3);
            Text text9 = new Text();
            text9.Text = "Sample Time";

            run9.Append(runProperties9);
            run9.Append(text9);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(run9);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph8);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "992", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge2 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders7 = new TableCellBorders();
            TopBorder topBorder7 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder8 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder8 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders7.Append(topBorder7);
            tableCellBorders7.Append(leftBorder8);
            tableCellBorders7.Append(rightBorder8);
            Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(verticalMerge2);
            tableCellProperties7.Append(tableCellBorders7);
            tableCellProperties7.Append(shading7);
            tableCellProperties7.Append(tableCellVerticalAlignment7);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "008162D3", RsidParagraphAddition = "008162D3", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            Justification justification6 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts18 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize5 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties9.Append(runFonts18);
            paragraphMarkRunProperties9.Append(fontSize5);
            paragraphMarkRunProperties9.Append(fontSizeComplexScript4);

            paragraphProperties9.Append(justification6);
            paragraphProperties9.Append(paragraphMarkRunProperties9);
            ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run10 = new Run();

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts19 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize6 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "20" };

            runProperties10.Append(runFonts19);
            runProperties10.Append(fontSize6);
            runProperties10.Append(fontSizeComplexScript5);
            Text text10 = new Text();
            DateTime? MinTime = labSheetA1Sheet.LabSheetA1MeasurementList.Where(c => c.Time != null).Min(c => c.Time);
            text10.Text = (MinTime == null ? "" : MinTime.Value.ToString("HH:mm"));

            run10.Append(runProperties10);
            run10.Append(text10);
            ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(proofError5);
            paragraph9.Append(run10);
            paragraph9.Append(proofError6);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            Justification justification7 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color2 = new Color() { Val = "C0C0C0" };
            FontSize fontSize7 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties10.Append(runFonts20);
            paragraphMarkRunProperties10.Append(color2);
            paragraphMarkRunProperties10.Append(fontSize7);
            paragraphMarkRunProperties10.Append(fontSizeComplexScript6);

            paragraphProperties10.Append(justification7);
            paragraphProperties10.Append(paragraphMarkRunProperties10);
            ProofError proofError7 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run11 = new Run();

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts21 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize8 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "20" };

            runProperties11.Append(runFonts21);
            runProperties11.Append(fontSize8);
            runProperties11.Append(fontSizeComplexScript7);
            Text text11 = new Text();
            DateTime? MaxTime = labSheetA1Sheet.LabSheetA1MeasurementList.Where(c => c.Time != null).Max(c => c.Time);
            text11.Text = (MaxTime == null ? "" : MaxTime.Value.ToString("HH:mm"));

            run11.Append(runProperties11);
            run11.Append(text11);
            ProofError proofError8 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(proofError7);
            paragraph10.Append(run11);
            paragraph10.Append(proofError8);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph9);
            tableCell7.Append(paragraph10);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "284", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge3 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders8 = new TableCellBorders();
            TopBorder topBorder8 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder9 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder9 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders8.Append(topBorder8);
            tableCellBorders8.Append(leftBorder9);
            tableCellBorders8.Append(rightBorder9);
            Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment8 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(verticalMerge3);
            tableCellProperties8.Append(tableCellBorders8);
            tableCellProperties8.Append(shading8);
            tableCellProperties8.Append(tableCellVerticalAlignment8);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts22 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize9 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties11.Append(runFonts22);
            paragraphMarkRunProperties11.Append(fontSize9);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript8);

            paragraphProperties11.Append(paragraphMarkRunProperties11);

            paragraph11.Append(paragraphProperties11);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph11);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "1417", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge4 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders9 = new TableCellBorders();
            TopBorder topBorder9 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder10 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder10 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders9.Append(topBorder9);
            tableCellBorders9.Append(leftBorder10);
            tableCellBorders9.Append(rightBorder10);
            Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment9 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(verticalMerge4);
            tableCellProperties9.Append(tableCellBorders9);
            tableCellProperties9.Append(shading9);
            tableCellProperties9.Append(tableCellVerticalAlignment9);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphAddition = "00D10A17", RsidParagraphProperties = "0032114E", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts23 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize10 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties12.Append(runFonts23);
            paragraphMarkRunProperties12.Append(fontSize10);
            paragraphMarkRunProperties12.Append(fontSizeComplexScript9);

            paragraphProperties12.Append(paragraphMarkRunProperties12);

            Run run12 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts24 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize11 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "20" };

            runProperties12.Append(runFonts24);
            runProperties12.Append(fontSize11);
            runProperties12.Append(fontSizeComplexScript10);
            Text text12 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text12.Text = "Incubation ";

            run12.Append(runProperties12);
            run12.Append(text12);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run12);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            Justification justification8 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            RunFonts runFonts25 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize12 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties13.Append(runFonts25);
            paragraphMarkRunProperties13.Append(fontSize12);
            paragraphMarkRunProperties13.Append(fontSizeComplexScript11);

            paragraphProperties13.Append(justification8);
            paragraphProperties13.Append(paragraphMarkRunProperties13);

            Run run13 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts26 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize13 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "20" };

            runProperties13.Append(runFonts26);
            runProperties13.Append(fontSize13);
            runProperties13.Append(fontSizeComplexScript12);
            Text text13 = new Text();
            text13.Text = "Start Time";

            run13.Append(runProperties13);
            run13.Append(text13);

            paragraph13.Append(paragraphProperties13);
            paragraph13.Append(run13);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph12);
            tableCell9.Append(paragraph13);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "1985", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge5 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders10 = new TableCellBorders();
            TopBorder topBorder10 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder11 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder11 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders10.Append(topBorder10);
            tableCellBorders10.Append(leftBorder11);
            tableCellBorders10.Append(rightBorder11);
            Shading shading10 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment10 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties10.Append(tableCellWidth10);
            tableCellProperties10.Append(verticalMerge5);
            tableCellProperties10.Append(tableCellBorders10);
            tableCellProperties10.Append(shading10);
            tableCellProperties10.Append(tableCellVerticalAlignment10);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            Justification justification9 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts27 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize14 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties14.Append(runFonts27);
            paragraphMarkRunProperties14.Append(fontSize14);
            paragraphMarkRunProperties14.Append(fontSizeComplexScript13);

            paragraphProperties14.Append(justification9);
            paragraphProperties14.Append(paragraphMarkRunProperties14);
            ProofError proofError9 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run14 = new Run();

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts28 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize15 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "20" };

            runProperties14.Append(runFonts28);
            runProperties14.Append(fontSize15);
            runProperties14.Append(fontSizeComplexScript14);
            Text text14 = new Text();
            text14.Text = labSheetA1Sheet.IncubationBath1StartTime +
                (labSheetA1Sheet.WaterBathCount > 1 ? "/" + labSheetA1Sheet.IncubationBath2StartTime : "") +
                (labSheetA1Sheet.WaterBathCount > 2 ? "/" + labSheetA1Sheet.IncubationBath3StartTime : "");

            run14.Append(runProperties14);
            run14.Append(text14);
            ProofError proofError10 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(proofError9);
            paragraph14.Append(run14);
            paragraph14.Append(proofError10);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph14);

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "283", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge6 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders11 = new TableCellBorders();
            TopBorder topBorder11 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder12 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder12 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders11.Append(topBorder11);
            tableCellBorders11.Append(leftBorder12);
            tableCellBorders11.Append(rightBorder12);
            Shading shading11 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment11 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties11.Append(tableCellWidth11);
            tableCellProperties11.Append(verticalMerge6);
            tableCellProperties11.Append(tableCellBorders11);
            tableCellProperties11.Append(shading11);
            tableCellProperties11.Append(tableCellVerticalAlignment11);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize16 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties15.Append(runFonts29);
            paragraphMarkRunProperties15.Append(fontSize16);
            paragraphMarkRunProperties15.Append(fontSizeComplexScript15);

            paragraphProperties15.Append(paragraphMarkRunProperties15);

            paragraph15.Append(paragraphProperties15);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph15);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "1416", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge7 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders12 = new TableCellBorders();
            TopBorder topBorder12 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder13 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder13 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders12.Append(topBorder12);
            tableCellBorders12.Append(leftBorder13);
            tableCellBorders12.Append(rightBorder13);
            Shading shading12 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment12 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties12.Append(tableCellWidth12);
            tableCellProperties12.Append(verticalMerge7);
            tableCellProperties12.Append(tableCellBorders12);
            tableCellProperties12.Append(shading12);
            tableCellProperties12.Append(tableCellVerticalAlignment12);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            Justification justification10 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            RunFonts runFonts30 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize17 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties16.Append(runFonts30);
            paragraphMarkRunProperties16.Append(fontSize17);
            paragraphMarkRunProperties16.Append(fontSizeComplexScript16);

            paragraphProperties16.Append(justification10);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            Run run15 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts31 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize18 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "20" };

            runProperties15.Append(runFonts31);
            runProperties15.Append(fontSize18);
            runProperties15.Append(fontSizeComplexScript17);
            Text text15 = new Text();
            text15.Text = "TC (";

            run15.Append(runProperties15);
            run15.Append(text15);
            ProofError proofError11 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run16 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts32 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize19 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "20" };
            VerticalTextAlignment verticalTextAlignment1 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties16.Append(runFonts32);
            runProperties16.Append(fontSize19);
            runProperties16.Append(fontSizeComplexScript18);
            runProperties16.Append(verticalTextAlignment1);
            Text text16 = new Text();
            text16.Text = "o";

            run16.Append(runProperties16);
            run16.Append(text16);

            Run run17 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts33 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize20 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "20" };

            runProperties17.Append(runFonts33);
            runProperties17.Append(fontSize20);
            runProperties17.Append(fontSizeComplexScript19);
            Text text17 = new Text();
            text17.Text = "C";

            run17.Append(runProperties17);
            run17.Append(text17);
            ProofError proofError12 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run18 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties18 = new RunProperties();
            RunFonts runFonts34 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize21 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "20" };

            runProperties18.Append(runFonts34);
            runProperties18.Append(fontSize21);
            runProperties18.Append(fontSizeComplexScript20);
            Text text18 = new Text();
            text18.Text = ")";

            run18.Append(runProperties18);
            run18.Append(text18);

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(run15);
            paragraph16.Append(proofError11);
            paragraph16.Append(run16);
            paragraph16.Append(run17);
            paragraph16.Append(proofError12);
            paragraph16.Append(run18);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph16);

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "238", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge8 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders13 = new TableCellBorders();
            TopBorder topBorder13 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder14 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder14 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders13.Append(topBorder13);
            tableCellBorders13.Append(leftBorder14);
            tableCellBorders13.Append(rightBorder14);
            Shading shading13 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment13 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties13.Append(tableCellWidth13);
            tableCellProperties13.Append(verticalMerge8);
            tableCellProperties13.Append(tableCellBorders13);
            tableCellProperties13.Append(shading13);
            tableCellProperties13.Append(tableCellVerticalAlignment13);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            RunFonts runFonts35 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize22 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties17.Append(runFonts35);
            paragraphMarkRunProperties17.Append(fontSize22);
            paragraphMarkRunProperties17.Append(fontSizeComplexScript21);

            paragraphProperties17.Append(paragraphMarkRunProperties17);

            paragraph17.Append(paragraphProperties17);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph17);

            TableCell tableCell14 = new TableCell();

            TableCellProperties tableCellProperties14 = new TableCellProperties();
            TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "1031", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders14 = new TableCellBorders();
            TopBorder topBorder14 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder15 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder15 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders14.Append(topBorder14);
            tableCellBorders14.Append(leftBorder15);
            tableCellBorders14.Append(bottomBorder6);
            tableCellBorders14.Append(rightBorder15);
            Shading shading14 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment14 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties14.Append(tableCellWidth14);
            tableCellProperties14.Append(tableCellBorders14);
            tableCellProperties14.Append(shading14);
            tableCellProperties14.Append(tableCellVerticalAlignment14);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunFonts runFonts36 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold3 = new Bold();
            FontSize fontSize23 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties18.Append(runFonts36);
            paragraphMarkRunProperties18.Append(bold3);
            paragraphMarkRunProperties18.Append(fontSize23);
            paragraphMarkRunProperties18.Append(fontSizeComplexScript22);

            paragraphProperties18.Append(paragraphMarkRunProperties18);

            Run run19 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties19 = new RunProperties();
            RunFonts runFonts37 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold4 = new Bold();
            FontSize fontSize24 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "20" };

            runProperties19.Append(runFonts37);
            runProperties19.Append(bold4);
            runProperties19.Append(fontSize24);
            runProperties19.Append(fontSizeComplexScript23);
            Text text19 = new Text();
            text19.Text = "Control";

            run19.Append(runProperties19);
            run19.Append(text19);

            paragraph18.Append(paragraphProperties18);
            paragraph18.Append(run19);

            tableCell14.Append(tableCellProperties14);
            tableCell14.Append(paragraph18);

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "906", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders15 = new TableCellBorders();
            TopBorder topBorder15 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder16 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder16 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders15.Append(topBorder15);
            tableCellBorders15.Append(leftBorder16);
            tableCellBorders15.Append(rightBorder16);
            Shading shading15 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment15 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties15.Append(tableCellWidth15);
            tableCellProperties15.Append(tableCellBorders15);
            tableCellProperties15.Append(shading15);
            tableCellProperties15.Append(tableCellVerticalAlignment15);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "006B4849", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0055454A", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts38 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic1 = new Italic();
            FontSize fontSize25 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties19.Append(runFonts38);
            paragraphMarkRunProperties19.Append(italic1);
            paragraphMarkRunProperties19.Append(fontSize25);
            paragraphMarkRunProperties19.Append(fontSizeComplexScript24);

            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run20 = new Run() { RsidRunProperties = "006B4849" };

            RunProperties runProperties20 = new RunProperties();
            RunFonts runFonts39 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic2 = new Italic();
            FontSize fontSize26 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "18" };

            runProperties20.Append(runFonts39);
            runProperties20.Append(italic2);
            runProperties20.Append(fontSize26);
            runProperties20.Append(fontSizeComplexScript25);
            Text text20 = new Text();
            text20.Text = "positive";

            run20.Append(runProperties20);
            run20.Append(text20);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run20);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph19);

            TableCell tableCell16 = new TableCell();

            TableCellProperties tableCellProperties16 = new TableCellProperties();
            TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "1087", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders16 = new TableCellBorders();
            TopBorder topBorder16 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder17 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder17 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders16.Append(topBorder16);
            tableCellBorders16.Append(leftBorder17);
            tableCellBorders16.Append(rightBorder17);
            Shading shading16 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment16 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties16.Append(tableCellWidth16);
            tableCellProperties16.Append(tableCellBorders16);
            tableCellProperties16.Append(shading16);
            tableCellProperties16.Append(tableCellVerticalAlignment16);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "006B4849", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0055454A", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            RunFonts runFonts40 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize27 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties20.Append(runFonts40);
            paragraphMarkRunProperties20.Append(fontSize27);
            paragraphMarkRunProperties20.Append(fontSizeComplexScript26);

            paragraphProperties20.Append(paragraphMarkRunProperties20);

            Run run21 = new Run() { RsidRunProperties = "006B4849" };

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts41 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic3 = new Italic();
            FontSize fontSize28 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "18" };

            runProperties21.Append(runFonts41);
            runProperties21.Append(italic3);
            runProperties21.Append(fontSize28);
            runProperties21.Append(fontSizeComplexScript27);
            Text text21 = new Text();
            text21.Text = "non target";

            run21.Append(runProperties21);
            run21.Append(text21);

            paragraph20.Append(paragraphProperties20);
            paragraph20.Append(run21);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph20);

            TableCell tableCell17 = new TableCell();

            TableCellProperties tableCellProperties17 = new TableCellProperties();
            TableCellWidth tableCellWidth17 = new TableCellWidth() { Width = "1007", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders17 = new TableCellBorders();
            TopBorder topBorder17 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder18 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder18 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders17.Append(topBorder17);
            tableCellBorders17.Append(leftBorder18);
            tableCellBorders17.Append(rightBorder18);
            Shading shading17 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment17 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties17.Append(tableCellWidth17);
            tableCellProperties17.Append(tableCellBorders17);
            tableCellProperties17.Append(shading17);
            tableCellProperties17.Append(tableCellVerticalAlignment17);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "006B4849", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0055454A", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts42 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic4 = new Italic();
            FontSize fontSize29 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties21.Append(runFonts42);
            paragraphMarkRunProperties21.Append(italic4);
            paragraphMarkRunProperties21.Append(fontSize29);
            paragraphMarkRunProperties21.Append(fontSizeComplexScript28);

            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run22 = new Run() { RsidRunProperties = "006B4849" };

            RunProperties runProperties22 = new RunProperties();
            RunFonts runFonts43 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic5 = new Italic();
            FontSize fontSize30 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "18" };

            runProperties22.Append(runFonts43);
            runProperties22.Append(italic5);
            runProperties22.Append(fontSize30);
            runProperties22.Append(fontSizeComplexScript29);
            Text text22 = new Text();
            text22.Text = "negative";

            run22.Append(runProperties22);
            run22.Append(text22);

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run22);

            tableCell17.Append(tableCellProperties17);
            tableCell17.Append(paragraph21);

            TableCell tableCell18 = new TableCell();

            TableCellProperties tableCellProperties18 = new TableCellProperties();
            TableCellWidth tableCellWidth18 = new TableCellWidth() { Width = "906", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge9 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders18 = new TableCellBorders();
            TopBorder topBorder18 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder19 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder19 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders18.Append(topBorder18);
            tableCellBorders18.Append(leftBorder19);
            tableCellBorders18.Append(rightBorder19);
            Shading shading18 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment18 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties18.Append(tableCellWidth18);
            tableCellProperties18.Append(verticalMerge9);
            tableCellProperties18.Append(tableCellBorders18);
            tableCellProperties18.Append(shading18);
            tableCellProperties18.Append(tableCellVerticalAlignment18);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            Justification justification11 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts44 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize31 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties22.Append(runFonts44);
            paragraphMarkRunProperties22.Append(fontSize31);
            paragraphMarkRunProperties22.Append(fontSizeComplexScript30);

            paragraphProperties22.Append(justification11);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            Run run23 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties23 = new RunProperties();
            RunFonts runFonts45 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize32 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "20" };

            runProperties23.Append(runFonts45);
            runProperties23.Append(fontSize32);
            runProperties23.Append(fontSizeComplexScript31);
            Text text23 = new Text();
            text23.Text = "Blank";

            run23.Append(runProperties23);
            run23.Append(text23);

            paragraph22.Append(paragraphProperties22);
            paragraph22.Append(run23);

            tableCell18.Append(tableCellProperties18);
            tableCell18.Append(paragraph22);

            TableCell tableCell19 = new TableCell();

            TableCellProperties tableCellProperties19 = new TableCellProperties();
            TableCellWidth tableCellWidth19 = new TableCellWidth() { Width = "906", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge10 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders19 = new TableCellBorders();
            TopBorder topBorder19 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder20 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder20 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders19.Append(topBorder19);
            tableCellBorders19.Append(leftBorder20);
            tableCellBorders19.Append(rightBorder20);
            Shading shading19 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment19 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties19.Append(tableCellWidth19);
            tableCellProperties19.Append(verticalMerge10);
            tableCellProperties19.Append(tableCellBorders19);
            tableCellProperties19.Append(shading19);
            tableCellProperties19.Append(tableCellVerticalAlignment19);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            Justification justification12 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts46 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize33 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties23.Append(runFonts46);
            paragraphMarkRunProperties23.Append(fontSize33);
            paragraphMarkRunProperties23.Append(fontSizeComplexScript32);

            paragraphProperties23.Append(justification12);
            paragraphProperties23.Append(paragraphMarkRunProperties23);

            Run run24 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties24 = new RunProperties();
            RunFonts runFonts47 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize34 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "20" };

            runProperties24.Append(runFonts47);
            runProperties24.Append(fontSize34);
            runProperties24.Append(fontSizeComplexScript33);
            Text text24 = new Text();
            text24.Text = "A1";

            run24.Append(runProperties24);
            run24.Append(text24);

            paragraph23.Append(paragraphProperties23);
            paragraph23.Append(run24);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            Justification justification13 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            RunFonts runFonts48 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize35 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties24.Append(runFonts48);
            paragraphMarkRunProperties24.Append(fontSize35);
            paragraphMarkRunProperties24.Append(fontSizeComplexScript34);

            paragraphProperties24.Append(justification13);
            paragraphProperties24.Append(paragraphMarkRunProperties24);

            Run run25 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties25 = new RunProperties();
            RunFonts runFonts49 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize36 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "20" };

            runProperties25.Append(runFonts49);
            runProperties25.Append(fontSize36);
            runProperties25.Append(fontSizeComplexScript35);
            Text text25 = new Text();
            text25.Text = "Media";

            run25.Append(runProperties25);
            run25.Append(text25);

            paragraph24.Append(paragraphProperties24);
            paragraph24.Append(run25);

            tableCell19.Append(tableCellProperties19);
            tableCell19.Append(paragraph23);
            tableCell19.Append(paragraph24);

            TableCell tableCell20 = new TableCell();

            TableCellProperties tableCellProperties20 = new TableCellProperties();
            TableCellWidth tableCellWidth20 = new TableCellWidth() { Width = "1631", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge11 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders20 = new TableCellBorders();
            TopBorder topBorder20 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder21 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder21 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders20.Append(topBorder20);
            tableCellBorders20.Append(leftBorder21);
            tableCellBorders20.Append(rightBorder21);
            Shading shading20 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment20 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties20.Append(tableCellWidth20);
            tableCellProperties20.Append(verticalMerge11);
            tableCellProperties20.Append(tableCellBorders20);
            tableCellProperties20.Append(shading20);
            tableCellProperties20.Append(tableCellVerticalAlignment20);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            Justification justification14 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            RunFonts runFonts50 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize37 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties25.Append(runFonts50);
            paragraphMarkRunProperties25.Append(fontSize37);
            paragraphMarkRunProperties25.Append(fontSizeComplexScript36);

            paragraphProperties25.Append(justification14);
            paragraphProperties25.Append(paragraphMarkRunProperties25);

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<w:smartTag w:uri=\"urn:schemas-microsoft-com:office:smarttags\" w:element=\"place\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r w:rsidRPr=\"00ED7FB2\"><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\" /><w:sz w:val=\"20\" /><w:szCs w:val=\"20\" /></w:rPr><w:t>Lot</w:t></w:r></w:smartTag>");

            Run run26 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties26 = new RunProperties();
            RunFonts runFonts51 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize38 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "20" };

            runProperties26.Append(runFonts51);
            runProperties26.Append(fontSize38);
            runProperties26.Append(fontSizeComplexScript37);
            Text text26 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text26.Text = " #";

            run26.Append(runProperties26);
            run26.Append(text26);

            Run run27 = new Run() { RsidRunProperties = "00ED7FB2", RsidRunAddition = "009A609A" };

            RunProperties runProperties27 = new RunProperties();
            RunFonts runFonts52 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize39 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "20" };

            runProperties27.Append(runFonts52);
            runProperties27.Append(fontSize39);
            runProperties27.Append(fontSizeComplexScript38);
            Text text27 = new Text();
            text27.Text = "’s";

            run27.Append(runProperties27);
            run27.Append(text27);

            paragraph25.Append(paragraphProperties25);
            paragraph25.Append(openXmlUnknownElement1);
            paragraph25.Append(run26);
            paragraph25.Append(run27);

            tableCell20.Append(tableCellProperties20);
            tableCell20.Append(paragraph25);

            tableRow2.Append(tableRowProperties1);
            tableRow2.Append(tableCell6);
            tableRow2.Append(tableCell7);
            tableRow2.Append(tableCell8);
            tableRow2.Append(tableCell9);
            tableRow2.Append(tableCell10);
            tableRow2.Append(tableCell11);
            tableRow2.Append(tableCell12);
            tableRow2.Append(tableCell13);
            tableRow2.Append(tableCell14);
            tableRow2.Append(tableCell15);
            tableRow2.Append(tableCell16);
            tableRow2.Append(tableCell17);
            tableRow2.Append(tableCell18);
            tableRow2.Append(tableCell19);
            tableRow2.Append(tableCell20);

            TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "003511D5", RsidTableRowProperties = "0032114E" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)230U };

            tableRowProperties2.Append(tableRowHeight2);

            TableCell tableCell21 = new TableCell();

            TableCellProperties tableCellProperties21 = new TableCellProperties();
            TableCellWidth tableCellWidth21 = new TableCellWidth() { Width = "1031", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge12 = new VerticalMerge();

            TableCellBorders tableCellBorders21 = new TableCellBorders();
            LeftBorder leftBorder22 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder7 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder22 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders21.Append(leftBorder22);
            tableCellBorders21.Append(bottomBorder7);
            tableCellBorders21.Append(rightBorder22);
            Shading shading21 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment21 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties21.Append(tableCellWidth21);
            tableCellProperties21.Append(verticalMerge12);
            tableCellProperties21.Append(tableCellBorders21);
            tableCellProperties21.Append(shading21);
            tableCellProperties21.Append(tableCellVerticalAlignment21);

            Paragraph paragraph26 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();
            Justification justification15 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
            RunFonts runFonts53 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize40 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties26.Append(runFonts53);
            paragraphMarkRunProperties26.Append(fontSize40);
            paragraphMarkRunProperties26.Append(fontSizeComplexScript39);

            paragraphProperties26.Append(justification15);
            paragraphProperties26.Append(paragraphMarkRunProperties26);

            paragraph26.Append(paragraphProperties26);

            tableCell21.Append(tableCellProperties21);
            tableCell21.Append(paragraph26);

            TableCell tableCell22 = new TableCell();

            TableCellProperties tableCellProperties22 = new TableCellProperties();
            TableCellWidth tableCellWidth22 = new TableCellWidth() { Width = "992", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge13 = new VerticalMerge();

            TableCellBorders tableCellBorders22 = new TableCellBorders();
            LeftBorder leftBorder23 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder8 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder23 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders22.Append(leftBorder23);
            tableCellBorders22.Append(bottomBorder8);
            tableCellBorders22.Append(rightBorder23);
            Shading shading22 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment22 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties22.Append(tableCellWidth22);
            tableCellProperties22.Append(verticalMerge13);
            tableCellProperties22.Append(tableCellBorders22);
            tableCellProperties22.Append(shading22);
            tableCellProperties22.Append(tableCellVerticalAlignment22);

            Paragraph paragraph27 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            Justification justification16 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties27 = new ParagraphMarkRunProperties();
            RunFonts runFonts54 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color3 = new Color() { Val = "C0C0C0" };
            FontSize fontSize41 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties27.Append(runFonts54);
            paragraphMarkRunProperties27.Append(color3);
            paragraphMarkRunProperties27.Append(fontSize41);
            paragraphMarkRunProperties27.Append(fontSizeComplexScript40);

            paragraphProperties27.Append(justification16);
            paragraphProperties27.Append(paragraphMarkRunProperties27);

            paragraph27.Append(paragraphProperties27);

            tableCell22.Append(tableCellProperties22);
            tableCell22.Append(paragraph27);

            TableCell tableCell23 = new TableCell();

            TableCellProperties tableCellProperties23 = new TableCellProperties();
            TableCellWidth tableCellWidth23 = new TableCellWidth() { Width = "284", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge14 = new VerticalMerge();

            TableCellBorders tableCellBorders23 = new TableCellBorders();
            LeftBorder leftBorder24 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder9 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder24 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders23.Append(leftBorder24);
            tableCellBorders23.Append(bottomBorder9);
            tableCellBorders23.Append(rightBorder24);
            Shading shading23 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment23 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties23.Append(tableCellWidth23);
            tableCellProperties23.Append(verticalMerge14);
            tableCellProperties23.Append(tableCellBorders23);
            tableCellProperties23.Append(shading23);
            tableCellProperties23.Append(tableCellVerticalAlignment23);

            Paragraph paragraph28 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            RunFonts runFonts55 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize42 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties28.Append(runFonts55);
            paragraphMarkRunProperties28.Append(fontSize42);
            paragraphMarkRunProperties28.Append(fontSizeComplexScript41);

            paragraphProperties28.Append(paragraphMarkRunProperties28);

            paragraph28.Append(paragraphProperties28);

            tableCell23.Append(tableCellProperties23);
            tableCell23.Append(paragraph28);

            TableCell tableCell24 = new TableCell();

            TableCellProperties tableCellProperties24 = new TableCellProperties();
            TableCellWidth tableCellWidth24 = new TableCellWidth() { Width = "1417", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge15 = new VerticalMerge();

            TableCellBorders tableCellBorders24 = new TableCellBorders();
            LeftBorder leftBorder25 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder10 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder25 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders24.Append(leftBorder25);
            tableCellBorders24.Append(bottomBorder10);
            tableCellBorders24.Append(rightBorder25);
            Shading shading24 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment24 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties24.Append(tableCellWidth24);
            tableCellProperties24.Append(verticalMerge15);
            tableCellProperties24.Append(tableCellBorders24);
            tableCellProperties24.Append(shading24);
            tableCellProperties24.Append(tableCellVerticalAlignment24);

            Paragraph paragraph29 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties29 = new ParagraphMarkRunProperties();
            RunFonts runFonts56 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize43 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties29.Append(runFonts56);
            paragraphMarkRunProperties29.Append(fontSize43);
            paragraphMarkRunProperties29.Append(fontSizeComplexScript42);

            paragraphProperties29.Append(paragraphMarkRunProperties29);

            paragraph29.Append(paragraphProperties29);

            tableCell24.Append(tableCellProperties24);
            tableCell24.Append(paragraph29);

            TableCell tableCell25 = new TableCell();

            TableCellProperties tableCellProperties25 = new TableCellProperties();
            TableCellWidth tableCellWidth25 = new TableCellWidth() { Width = "1985", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge16 = new VerticalMerge();

            TableCellBorders tableCellBorders25 = new TableCellBorders();
            LeftBorder leftBorder26 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder11 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder26 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders25.Append(leftBorder26);
            tableCellBorders25.Append(bottomBorder11);
            tableCellBorders25.Append(rightBorder26);
            Shading shading25 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment25 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties25.Append(tableCellWidth25);
            tableCellProperties25.Append(verticalMerge16);
            tableCellProperties25.Append(tableCellBorders25);
            tableCellProperties25.Append(shading25);
            tableCellProperties25.Append(tableCellVerticalAlignment25);

            Paragraph paragraph30 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();
            Justification justification17 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties30 = new ParagraphMarkRunProperties();
            RunFonts runFonts57 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize44 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties30.Append(runFonts57);
            paragraphMarkRunProperties30.Append(fontSize44);
            paragraphMarkRunProperties30.Append(fontSizeComplexScript43);

            paragraphProperties30.Append(justification17);
            paragraphProperties30.Append(paragraphMarkRunProperties30);

            paragraph30.Append(paragraphProperties30);

            tableCell25.Append(tableCellProperties25);
            tableCell25.Append(paragraph30);

            TableCell tableCell26 = new TableCell();

            TableCellProperties tableCellProperties26 = new TableCellProperties();
            TableCellWidth tableCellWidth26 = new TableCellWidth() { Width = "283", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge17 = new VerticalMerge();

            TableCellBorders tableCellBorders26 = new TableCellBorders();
            LeftBorder leftBorder27 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder12 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder27 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders26.Append(leftBorder27);
            tableCellBorders26.Append(bottomBorder12);
            tableCellBorders26.Append(rightBorder27);
            Shading shading26 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment26 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties26.Append(tableCellWidth26);
            tableCellProperties26.Append(verticalMerge17);
            tableCellProperties26.Append(tableCellBorders26);
            tableCellProperties26.Append(shading26);
            tableCellProperties26.Append(tableCellVerticalAlignment26);

            Paragraph paragraph31 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties31 = new ParagraphMarkRunProperties();
            RunFonts runFonts58 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize45 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties31.Append(runFonts58);
            paragraphMarkRunProperties31.Append(fontSize45);
            paragraphMarkRunProperties31.Append(fontSizeComplexScript44);

            paragraphProperties31.Append(paragraphMarkRunProperties31);

            paragraph31.Append(paragraphProperties31);

            tableCell26.Append(tableCellProperties26);
            tableCell26.Append(paragraph31);

            TableCell tableCell27 = new TableCell();

            TableCellProperties tableCellProperties27 = new TableCellProperties();
            TableCellWidth tableCellWidth27 = new TableCellWidth() { Width = "1416", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge18 = new VerticalMerge();

            TableCellBorders tableCellBorders27 = new TableCellBorders();
            LeftBorder leftBorder28 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder13 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder28 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders27.Append(leftBorder28);
            tableCellBorders27.Append(bottomBorder13);
            tableCellBorders27.Append(rightBorder28);
            Shading shading27 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment27 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties27.Append(tableCellWidth27);
            tableCellProperties27.Append(verticalMerge18);
            tableCellProperties27.Append(tableCellBorders27);
            tableCellProperties27.Append(shading27);
            tableCellProperties27.Append(tableCellVerticalAlignment27);

            Paragraph paragraph32 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties32 = new ParagraphMarkRunProperties();
            RunFonts runFonts59 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize46 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties32.Append(runFonts59);
            paragraphMarkRunProperties32.Append(fontSize46);
            paragraphMarkRunProperties32.Append(fontSizeComplexScript45);

            paragraphProperties32.Append(paragraphMarkRunProperties32);

            paragraph32.Append(paragraphProperties32);

            tableCell27.Append(tableCellProperties27);
            tableCell27.Append(paragraph32);

            TableCell tableCell28 = new TableCell();

            TableCellProperties tableCellProperties28 = new TableCellProperties();
            TableCellWidth tableCellWidth28 = new TableCellWidth() { Width = "238", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge19 = new VerticalMerge();

            TableCellBorders tableCellBorders28 = new TableCellBorders();
            LeftBorder leftBorder29 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder14 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder29 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders28.Append(leftBorder29);
            tableCellBorders28.Append(bottomBorder14);
            tableCellBorders28.Append(rightBorder29);
            Shading shading28 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment28 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties28.Append(tableCellWidth28);
            tableCellProperties28.Append(verticalMerge19);
            tableCellProperties28.Append(tableCellBorders28);
            tableCellProperties28.Append(shading28);
            tableCellProperties28.Append(tableCellVerticalAlignment28);

            Paragraph paragraph33 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties33 = new ParagraphMarkRunProperties();
            RunFonts runFonts60 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize47 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties33.Append(runFonts60);
            paragraphMarkRunProperties33.Append(fontSize47);
            paragraphMarkRunProperties33.Append(fontSizeComplexScript46);

            paragraphProperties33.Append(paragraphMarkRunProperties33);

            paragraph33.Append(paragraphProperties33);

            tableCell28.Append(tableCellProperties28);
            tableCell28.Append(paragraph33);

            TableCell tableCell29 = new TableCell();

            TableCellProperties tableCellProperties29 = new TableCellProperties();
            TableCellWidth tableCellWidth29 = new TableCellWidth() { Width = "1031", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders29 = new TableCellBorders();
            TopBorder topBorder21 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder30 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder15 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder30 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders29.Append(topBorder21);
            tableCellBorders29.Append(leftBorder30);
            tableCellBorders29.Append(bottomBorder15);
            tableCellBorders29.Append(rightBorder30);
            Shading shading29 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment29 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties29.Append(tableCellWidth29);
            tableCellProperties29.Append(tableCellBorders29);
            tableCellProperties29.Append(shading29);
            tableCellProperties29.Append(tableCellVerticalAlignment29);

            Paragraph paragraph34 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties34 = new ParagraphMarkRunProperties();
            RunFonts runFonts61 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold5 = new Bold();
            FontSize fontSize48 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties34.Append(runFonts61);
            paragraphMarkRunProperties34.Append(bold5);
            paragraphMarkRunProperties34.Append(fontSize48);
            paragraphMarkRunProperties34.Append(fontSizeComplexScript47);

            paragraphProperties34.Append(paragraphMarkRunProperties34);

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<w:smartTag w:uri=\"urn:schemas-microsoft-com:office:smarttags\" w:element=\"place\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r w:rsidRPr=\"00ED7FB2\"><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\" /><w:b /><w:sz w:val=\"20\" /><w:szCs w:val=\"20\" /></w:rPr><w:t>Lot</w:t></w:r></w:smartTag>");

            Run run28 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties28 = new RunProperties();
            RunFonts runFonts62 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold6 = new Bold();
            FontSize fontSize49 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "20" };

            runProperties28.Append(runFonts62);
            runProperties28.Append(bold6);
            runProperties28.Append(fontSize49);
            runProperties28.Append(fontSizeComplexScript48);
            Text text28 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text28.Text = " #’s";

            run28.Append(runProperties28);
            run28.Append(text28);

            paragraph34.Append(paragraphProperties34);
            paragraph34.Append(openXmlUnknownElement2);
            paragraph34.Append(run28);

            tableCell29.Append(tableCellProperties29);
            tableCell29.Append(paragraph34);

            TableCell tableCell30 = new TableCell();

            TableCellProperties tableCellProperties30 = new TableCellProperties();
            TableCellWidth tableCellWidth30 = new TableCellWidth() { Width = "3000", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = 3 };

            TableCellBorders tableCellBorders30 = new TableCellBorders();
            LeftBorder leftBorder31 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder31 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders30.Append(leftBorder31);
            tableCellBorders30.Append(rightBorder31);
            Shading shading30 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment30 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties30.Append(tableCellWidth30);
            tableCellProperties30.Append(gridSpan1);
            tableCellProperties30.Append(tableCellBorders30);
            tableCellProperties30.Append(shading30);
            tableCellProperties30.Append(tableCellVerticalAlignment30);

            Paragraph paragraph35 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            RunFonts runFonts63 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic6 = new Italic();
            FontSize fontSize50 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties35.Append(runFonts63);
            paragraphMarkRunProperties35.Append(italic6);
            paragraphMarkRunProperties35.Append(fontSize50);
            paragraphMarkRunProperties35.Append(fontSizeComplexScript49);

            paragraphProperties35.Append(paragraphMarkRunProperties35);

            Run run29 = new Run();

            RunProperties runProperties29 = new RunProperties();
            RunFonts runFonts64 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize51 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "20" };

            runProperties29.Append(runFonts64);
            runProperties29.Append(fontSize51);
            runProperties29.Append(fontSizeComplexScript50);
            Text text29 = new Text();
            text29.Text = (labSheetA1Sheet.ControlLot.Length > 25 ? labSheetA1Sheet.ControlLot.Substring(0, 25) + " ..." : labSheetA1Sheet.ControlLot);

            run29.Append(runProperties29);
            run29.Append(text29);

            paragraph35.Append(paragraphProperties35);
            paragraph35.Append(run29);

            tableCell30.Append(tableCellProperties30);
            tableCell30.Append(paragraph35);

            TableCell tableCell31 = new TableCell();

            TableCellProperties tableCellProperties31 = new TableCellProperties();
            TableCellWidth tableCellWidth31 = new TableCellWidth() { Width = "906", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge20 = new VerticalMerge();

            TableCellBorders tableCellBorders31 = new TableCellBorders();
            LeftBorder leftBorder32 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder16 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder32 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders31.Append(leftBorder32);
            tableCellBorders31.Append(bottomBorder16);
            tableCellBorders31.Append(rightBorder32);
            Shading shading31 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment31 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties31.Append(tableCellWidth31);
            tableCellProperties31.Append(verticalMerge20);
            tableCellProperties31.Append(tableCellBorders31);
            tableCellProperties31.Append(shading31);
            tableCellProperties31.Append(tableCellVerticalAlignment31);

            Paragraph paragraph36 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();
            Justification justification18 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            RunFonts runFonts65 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize52 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties36.Append(runFonts65);
            paragraphMarkRunProperties36.Append(fontSize52);
            paragraphMarkRunProperties36.Append(fontSizeComplexScript51);

            paragraphProperties36.Append(justification18);
            paragraphProperties36.Append(paragraphMarkRunProperties36);

            paragraph36.Append(paragraphProperties36);

            tableCell31.Append(tableCellProperties31);
            tableCell31.Append(paragraph36);

            TableCell tableCell32 = new TableCell();

            TableCellProperties tableCellProperties32 = new TableCellProperties();
            TableCellWidth tableCellWidth32 = new TableCellWidth() { Width = "906", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge21 = new VerticalMerge();

            TableCellBorders tableCellBorders32 = new TableCellBorders();
            LeftBorder leftBorder33 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder17 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder33 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders32.Append(leftBorder33);
            tableCellBorders32.Append(bottomBorder17);
            tableCellBorders32.Append(rightBorder33);
            Shading shading32 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment32 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties32.Append(tableCellWidth32);
            tableCellProperties32.Append(verticalMerge21);
            tableCellProperties32.Append(tableCellBorders32);
            tableCellProperties32.Append(shading32);
            tableCellProperties32.Append(tableCellVerticalAlignment32);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();
            Justification justification19 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            RunFonts runFonts66 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize53 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties37.Append(runFonts66);
            paragraphMarkRunProperties37.Append(fontSize53);
            paragraphMarkRunProperties37.Append(fontSizeComplexScript52);

            paragraphProperties37.Append(justification19);
            paragraphProperties37.Append(paragraphMarkRunProperties37);

            paragraph37.Append(paragraphProperties37);

            tableCell32.Append(tableCellProperties32);
            tableCell32.Append(paragraph37);

            TableCell tableCell33 = new TableCell();

            TableCellProperties tableCellProperties33 = new TableCellProperties();
            TableCellWidth tableCellWidth33 = new TableCellWidth() { Width = "1631", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge22 = new VerticalMerge();

            TableCellBorders tableCellBorders33 = new TableCellBorders();
            LeftBorder leftBorder34 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder18 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder34 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders33.Append(leftBorder34);
            tableCellBorders33.Append(bottomBorder18);
            tableCellBorders33.Append(rightBorder34);
            Shading shading33 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment33 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties33.Append(tableCellWidth33);
            tableCellProperties33.Append(verticalMerge22);
            tableCellProperties33.Append(tableCellBorders33);
            tableCellProperties33.Append(shading33);
            tableCellProperties33.Append(tableCellVerticalAlignment33);

            Paragraph paragraph38 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "003511D5", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "003511D5" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();
            Justification justification20 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties38 = new ParagraphMarkRunProperties();
            RunFonts runFonts67 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize54 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties38.Append(runFonts67);
            paragraphMarkRunProperties38.Append(fontSize54);
            paragraphMarkRunProperties38.Append(fontSizeComplexScript53);

            paragraphProperties38.Append(justification20);
            paragraphProperties38.Append(paragraphMarkRunProperties38);

            paragraph38.Append(paragraphProperties38);

            tableCell33.Append(tableCellProperties33);
            tableCell33.Append(paragraph38);

            tableRow3.Append(tableRowProperties2);
            tableRow3.Append(tableCell21);
            tableRow3.Append(tableCell22);
            tableRow3.Append(tableCell23);
            tableRow3.Append(tableCell24);
            tableRow3.Append(tableCell25);
            tableRow3.Append(tableCell26);
            tableRow3.Append(tableCell27);
            tableRow3.Append(tableCell28);
            tableRow3.Append(tableCell29);
            tableRow3.Append(tableCell30);
            tableRow3.Append(tableCell31);
            tableRow3.Append(tableCell32);
            tableRow3.Append(tableCell33);

            TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00397699", RsidTableRowProperties = "0032114E" };

            TableCell tableCell34 = new TableCell();

            TableCellProperties tableCellProperties34 = new TableCellProperties();
            TableCellWidth tableCellWidth34 = new TableCellWidth() { Width = "1031", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders34 = new TableCellBorders();
            TopBorder topBorder22 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder35 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder19 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder35 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders34.Append(topBorder22);
            tableCellBorders34.Append(leftBorder35);
            tableCellBorders34.Append(bottomBorder19);
            tableCellBorders34.Append(rightBorder35);
            Shading shading34 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment34 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties34.Append(tableCellWidth34);
            tableCellProperties34.Append(tableCellBorders34);
            tableCellProperties34.Append(shading34);
            tableCellProperties34.Append(tableCellVerticalAlignment34);

            Paragraph paragraph39 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();
            Justification justification21 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties39 = new ParagraphMarkRunProperties();
            RunFonts runFonts68 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize55 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties39.Append(runFonts68);
            paragraphMarkRunProperties39.Append(fontSize55);
            paragraphMarkRunProperties39.Append(fontSizeComplexScript54);

            paragraphProperties39.Append(justification21);
            paragraphProperties39.Append(paragraphMarkRunProperties39);

            Run run30 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties30 = new RunProperties();
            RunFonts runFonts69 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize56 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "20" };

            runProperties30.Append(runFonts69);
            runProperties30.Append(fontSize56);
            runProperties30.Append(fontSizeComplexScript55);
            Text text30 = new Text();
            text30.Text = "Tide";

            run30.Append(runProperties30);
            run30.Append(text30);

            paragraph39.Append(paragraphProperties39);
            paragraph39.Append(run30);

            tableCell34.Append(tableCellProperties34);
            tableCell34.Append(paragraph39);

            TableCell tableCell35 = new TableCell();

            TableCellProperties tableCellProperties35 = new TableCellProperties();
            TableCellWidth tableCellWidth35 = new TableCellWidth() { Width = "992", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders35 = new TableCellBorders();
            TopBorder topBorder23 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder36 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder20 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder36 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders35.Append(topBorder23);
            tableCellBorders35.Append(leftBorder36);
            tableCellBorders35.Append(bottomBorder20);
            tableCellBorders35.Append(rightBorder36);
            Shading shading35 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment35 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties35.Append(tableCellWidth35);
            tableCellProperties35.Append(tableCellBorders35);
            tableCellProperties35.Append(shading35);
            tableCellProperties35.Append(tableCellVerticalAlignment35);

            Paragraph paragraph40 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties40 = new ParagraphProperties();
            Justification justification22 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties40 = new ParagraphMarkRunProperties();
            RunFonts runFonts70 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize57 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties40.Append(runFonts70);
            paragraphMarkRunProperties40.Append(fontSize57);
            paragraphMarkRunProperties40.Append(fontSizeComplexScript56);

            paragraphProperties40.Append(justification22);
            paragraphProperties40.Append(paragraphMarkRunProperties40);
            ProofError proofError13 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run31 = new Run();

            RunProperties runProperties31 = new RunProperties();
            RunFonts runFonts71 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize58 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript57 = new FontSizeComplexScript() { Val = "20" };

            runProperties31.Append(runFonts71);
            runProperties31.Append(fontSize58);
            runProperties31.Append(fontSizeComplexScript57);
            Text text31 = new Text();
            text31.Text = "";
            if (!labSheetA1Sheet.Tides.Contains("-"))
            {
                text31.Text = labSheetA1Sheet.Tides;
            }

            run31.Append(runProperties31);
            run31.Append(text31);
            ProofError proofError14 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph40.Append(paragraphProperties40);
            paragraph40.Append(proofError13);
            paragraph40.Append(run31);
            paragraph40.Append(proofError14);

            tableCell35.Append(tableCellProperties35);
            tableCell35.Append(paragraph40);

            TableCell tableCell36 = new TableCell();

            TableCellProperties tableCellProperties36 = new TableCellProperties();
            TableCellWidth tableCellWidth36 = new TableCellWidth() { Width = "284", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders36 = new TableCellBorders();
            TopBorder topBorder24 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder37 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder21 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder37 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders36.Append(topBorder24);
            tableCellBorders36.Append(leftBorder37);
            tableCellBorders36.Append(bottomBorder21);
            tableCellBorders36.Append(rightBorder37);
            Shading shading36 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment36 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties36.Append(tableCellWidth36);
            tableCellProperties36.Append(tableCellBorders36);
            tableCellProperties36.Append(shading36);
            tableCellProperties36.Append(tableCellVerticalAlignment36);

            Paragraph paragraph41 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties41 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties41 = new ParagraphMarkRunProperties();
            RunFonts runFonts72 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize59 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript58 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties41.Append(runFonts72);
            paragraphMarkRunProperties41.Append(fontSize59);
            paragraphMarkRunProperties41.Append(fontSizeComplexScript58);

            paragraphProperties41.Append(paragraphMarkRunProperties41);

            paragraph41.Append(paragraphProperties41);

            tableCell36.Append(tableCellProperties36);
            tableCell36.Append(paragraph41);

            TableCell tableCell37 = new TableCell();

            TableCellProperties tableCellProperties37 = new TableCellProperties();
            TableCellWidth tableCellWidth37 = new TableCellWidth() { Width = "1417", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders37 = new TableCellBorders();
            TopBorder topBorder25 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder38 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder22 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder38 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders37.Append(topBorder25);
            tableCellBorders37.Append(leftBorder38);
            tableCellBorders37.Append(bottomBorder22);
            tableCellBorders37.Append(rightBorder38);
            Shading shading37 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment37 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties37.Append(tableCellWidth37);
            tableCellProperties37.Append(tableCellBorders37);
            tableCellProperties37.Append(shading37);
            tableCellProperties37.Append(tableCellVerticalAlignment37);

            Paragraph paragraph42 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties42 = new ParagraphProperties();
            Justification justification23 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties42 = new ParagraphMarkRunProperties();
            RunFonts runFonts73 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize60 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript59 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties42.Append(runFonts73);
            paragraphMarkRunProperties42.Append(fontSize60);
            paragraphMarkRunProperties42.Append(fontSizeComplexScript59);

            paragraphProperties42.Append(justification23);
            paragraphProperties42.Append(paragraphMarkRunProperties42);

            Run run32 = new Run() { RsidRunProperties = "006B4849" };

            RunProperties runProperties32 = new RunProperties();
            RunFonts runFonts74 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize61 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript60 = new FontSizeComplexScript() { Val = "20" };
            Highlight highlight2 = new Highlight() { Val = HighlightColorValues.Yellow };

            runProperties32.Append(runFonts74);
            runProperties32.Append(fontSize61);
            runProperties32.Append(fontSizeComplexScript60);
            runProperties32.Append(highlight2);
            Text text32 = new Text();
            text32.Text = "End";

            run32.Append(runProperties32);
            run32.Append(text32);

            Run run33 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties33 = new RunProperties();
            RunFonts runFonts75 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize62 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript61 = new FontSizeComplexScript() { Val = "20" };

            runProperties33.Append(runFonts75);
            runProperties33.Append(fontSize62);
            runProperties33.Append(fontSizeComplexScript61);
            Text text33 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text33.Text = " Time";

            run33.Append(runProperties33);
            run33.Append(text33);

            paragraph42.Append(paragraphProperties42);
            paragraph42.Append(run32);
            paragraph42.Append(run33);

            tableCell37.Append(tableCellProperties37);
            tableCell37.Append(paragraph42);

            TableCell tableCell38 = new TableCell();

            TableCellProperties tableCellProperties38 = new TableCellProperties();
            TableCellWidth tableCellWidth38 = new TableCellWidth() { Width = "1985", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders38 = new TableCellBorders();
            TopBorder topBorder26 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder39 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder23 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder39 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders38.Append(topBorder26);
            tableCellBorders38.Append(leftBorder39);
            tableCellBorders38.Append(bottomBorder23);
            tableCellBorders38.Append(rightBorder39);
            Shading shading38 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment38 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties38.Append(tableCellWidth38);
            tableCellProperties38.Append(tableCellBorders38);
            tableCellProperties38.Append(shading38);
            tableCellProperties38.Append(tableCellVerticalAlignment38);

            Paragraph paragraph43 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties43 = new ParagraphProperties();
            Justification justification24 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties43 = new ParagraphMarkRunProperties();
            RunFonts runFonts76 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize63 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript62 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties43.Append(runFonts76);
            paragraphMarkRunProperties43.Append(fontSize63);
            paragraphMarkRunProperties43.Append(fontSizeComplexScript62);

            paragraphProperties43.Append(justification24);
            paragraphProperties43.Append(paragraphMarkRunProperties43);

            Run run34 = new Run();

            RunProperties runProperties34 = new RunProperties();
            RunFonts runFonts77 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize64 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript63 = new FontSizeComplexScript() { Val = "20" };

            runProperties34.Append(runFonts77);
            runProperties34.Append(fontSize64);
            runProperties34.Append(fontSizeComplexScript63);
            Text text34 = new Text();
            text34.Text = labSheetA1Sheet.IncubationBath1EndTime +
                (labSheetA1Sheet.WaterBathCount > 1 ? "/" + labSheetA1Sheet.IncubationBath2EndTime : "") +
                (labSheetA1Sheet.WaterBathCount > 2 ? "/" + labSheetA1Sheet.IncubationBath3EndTime : "");

            run34.Append(runProperties34);
            run34.Append(text34);

            paragraph43.Append(paragraphProperties43);
            paragraph43.Append(run34);

            tableCell38.Append(tableCellProperties38);
            tableCell38.Append(paragraph43);

            TableCell tableCell39 = new TableCell();

            TableCellProperties tableCellProperties39 = new TableCellProperties();
            TableCellWidth tableCellWidth39 = new TableCellWidth() { Width = "283", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders39 = new TableCellBorders();
            TopBorder topBorder27 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder40 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder24 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder40 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders39.Append(topBorder27);
            tableCellBorders39.Append(leftBorder40);
            tableCellBorders39.Append(bottomBorder24);
            tableCellBorders39.Append(rightBorder40);
            Shading shading39 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment39 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties39.Append(tableCellWidth39);
            tableCellProperties39.Append(tableCellBorders39);
            tableCellProperties39.Append(shading39);
            tableCellProperties39.Append(tableCellVerticalAlignment39);

            Paragraph paragraph44 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties44 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties44 = new ParagraphMarkRunProperties();
            RunFonts runFonts78 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color4 = new Color() { Val = "DDDDDD" };
            FontSize fontSize65 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript64 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties44.Append(runFonts78);
            paragraphMarkRunProperties44.Append(color4);
            paragraphMarkRunProperties44.Append(fontSize65);
            paragraphMarkRunProperties44.Append(fontSizeComplexScript64);

            paragraphProperties44.Append(paragraphMarkRunProperties44);

            paragraph44.Append(paragraphProperties44);

            tableCell39.Append(tableCellProperties39);
            tableCell39.Append(paragraph44);

            TableCell tableCell40 = new TableCell();

            TableCellProperties tableCellProperties40 = new TableCellProperties();
            TableCellWidth tableCellWidth40 = new TableCellWidth() { Width = "1416", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders40 = new TableCellBorders();
            TopBorder topBorder28 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder41 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder25 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder41 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders40.Append(topBorder28);
            tableCellBorders40.Append(leftBorder41);
            tableCellBorders40.Append(bottomBorder25);
            tableCellBorders40.Append(rightBorder41);
            Shading shading40 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment40 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties40.Append(tableCellWidth40);
            tableCellProperties40.Append(tableCellBorders40);
            tableCellProperties40.Append(shading40);
            tableCellProperties40.Append(tableCellVerticalAlignment40);

            Paragraph paragraph45 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties45 = new ParagraphProperties();
            Justification justification25 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties45 = new ParagraphMarkRunProperties();
            RunFonts runFonts79 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color5 = new Color() { Val = "C0C0C0" };
            FontSize fontSize66 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript65 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties45.Append(runFonts79);
            paragraphMarkRunProperties45.Append(color5);
            paragraphMarkRunProperties45.Append(fontSize66);
            paragraphMarkRunProperties45.Append(fontSizeComplexScript65);

            paragraphProperties45.Append(justification25);
            paragraphProperties45.Append(paragraphMarkRunProperties45);
            ProofError proofError15 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run35 = new Run();

            RunProperties runProperties35 = new RunProperties();
            RunFonts runFonts80 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize67 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript66 = new FontSizeComplexScript() { Val = "20" };

            runProperties35.Append(runFonts80);
            runProperties35.Append(fontSize67);
            runProperties35.Append(fontSizeComplexScript66);
            Text text35 = new Text();
            text35.Text = labSheetA1Sheet.TCField1 + (labSheetA1Sheet.TCField2 != null && labSheetA1Sheet.TCField2.Length > 0 ? " / " + labSheetA1Sheet.TCField2 : "");

            run35.Append(runProperties35);
            run35.Append(text35);
            ProofError proofError16 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph45.Append(paragraphProperties45);
            paragraph45.Append(proofError15);
            paragraph45.Append(run35);
            paragraph45.Append(proofError16);

            tableCell40.Append(tableCellProperties40);
            tableCell40.Append(paragraph45);

            TableCell tableCell41 = new TableCell();

            TableCellProperties tableCellProperties41 = new TableCellProperties();
            TableCellWidth tableCellWidth41 = new TableCellWidth() { Width = "238", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders41 = new TableCellBorders();
            TopBorder topBorder29 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder42 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder26 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder42 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders41.Append(topBorder29);
            tableCellBorders41.Append(leftBorder42);
            tableCellBorders41.Append(bottomBorder26);
            tableCellBorders41.Append(rightBorder42);
            Shading shading41 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment41 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties41.Append(tableCellWidth41);
            tableCellProperties41.Append(tableCellBorders41);
            tableCellProperties41.Append(shading41);
            tableCellProperties41.Append(tableCellVerticalAlignment41);

            Paragraph paragraph46 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties46 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties46 = new ParagraphMarkRunProperties();
            RunFonts runFonts81 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic7 = new Italic();
            FontSize fontSize68 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript67 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties46.Append(runFonts81);
            paragraphMarkRunProperties46.Append(italic7);
            paragraphMarkRunProperties46.Append(fontSize68);
            paragraphMarkRunProperties46.Append(fontSizeComplexScript67);

            paragraphProperties46.Append(paragraphMarkRunProperties46);

            paragraph46.Append(paragraphProperties46);

            tableCell41.Append(tableCellProperties41);
            tableCell41.Append(paragraph46);

            TableCell tableCell42 = new TableCell();

            TableCellProperties tableCellProperties42 = new TableCellProperties();
            TableCellWidth tableCellWidth42 = new TableCellWidth() { Width = "1031", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders42 = new TableCellBorders();
            TopBorder topBorder30 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder43 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder27 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder43 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders42.Append(topBorder30);
            tableCellBorders42.Append(leftBorder43);
            tableCellBorders42.Append(bottomBorder27);
            tableCellBorders42.Append(rightBorder43);
            Shading shading42 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment42 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties42.Append(tableCellWidth42);
            tableCellProperties42.Append(tableCellBorders42);
            tableCellProperties42.Append(shading42);
            tableCellProperties42.Append(tableCellVerticalAlignment42);

            Paragraph paragraph47 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0055454A", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties47 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties47 = new ParagraphMarkRunProperties();
            RunFonts runFonts82 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize69 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript68 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties47.Append(runFonts82);
            paragraphMarkRunProperties47.Append(fontSize69);
            paragraphMarkRunProperties47.Append(fontSizeComplexScript68);

            paragraphProperties47.Append(paragraphMarkRunProperties47);

            Run run36 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties36 = new RunProperties();
            RunFonts runFonts83 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize70 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript69 = new FontSizeComplexScript() { Val = "20" };

            runProperties36.Append(runFonts83);
            runProperties36.Append(fontSize70);
            runProperties36.Append(fontSizeComplexScript69);
            Text text36 = new Text();
            text36.Text = "35.0";

            run36.Append(runProperties36);
            run36.Append(text36);

            Run run37 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties37 = new RunProperties();
            RunFonts runFonts84 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize71 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript70 = new FontSizeComplexScript() { Val = "20" };
            VerticalTextAlignment verticalTextAlignment2 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties37.Append(runFonts84);
            runProperties37.Append(fontSize71);
            runProperties37.Append(fontSizeComplexScript70);
            runProperties37.Append(verticalTextAlignment2);
            Text text37 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text37.Text = " ";

            run37.Append(runProperties37);
            run37.Append(text37);
            ProofError proofError17 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run38 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties38 = new RunProperties();
            RunFonts runFonts85 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize72 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript71 = new FontSizeComplexScript() { Val = "20" };
            VerticalTextAlignment verticalTextAlignment3 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties38.Append(runFonts85);
            runProperties38.Append(fontSize72);
            runProperties38.Append(fontSizeComplexScript71);
            runProperties38.Append(verticalTextAlignment3);
            Text text38 = new Text();
            text38.Text = "o";

            run38.Append(runProperties38);
            run38.Append(text38);

            Run run39 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties39 = new RunProperties();
            RunFonts runFonts86 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize73 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript72 = new FontSizeComplexScript() { Val = "20" };

            runProperties39.Append(runFonts86);
            runProperties39.Append(fontSize73);
            runProperties39.Append(fontSizeComplexScript72);
            Text text39 = new Text();
            text39.Text = "C";

            run39.Append(runProperties39);
            run39.Append(text39);
            ProofError proofError18 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph47.Append(paragraphProperties47);
            paragraph47.Append(run36);
            paragraph47.Append(run37);
            paragraph47.Append(proofError17);
            paragraph47.Append(run38);
            paragraph47.Append(run39);
            paragraph47.Append(proofError18);

            tableCell42.Append(tableCellProperties42);
            tableCell42.Append(paragraph47);

            TableCell tableCell43 = new TableCell();

            TableCellProperties tableCellProperties43 = new TableCellProperties();
            TableCellWidth tableCellWidth43 = new TableCellWidth() { Width = "906", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders43 = new TableCellBorders();
            LeftBorder leftBorder44 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder44 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders43.Append(leftBorder44);
            tableCellBorders43.Append(rightBorder44);
            Shading shading43 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment43 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties43.Append(tableCellWidth43);
            tableCellProperties43.Append(tableCellBorders43);
            tableCellProperties43.Append(shading43);
            tableCellProperties43.Append(tableCellVerticalAlignment43);

            Paragraph paragraph48 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties48 = new ParagraphProperties();
            Justification justification26 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties48 = new ParagraphMarkRunProperties();
            RunFonts runFonts87 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic8 = new Italic();
            FontSize fontSize74 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript73 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties48.Append(runFonts87);
            paragraphMarkRunProperties48.Append(italic8);
            paragraphMarkRunProperties48.Append(fontSize74);
            paragraphMarkRunProperties48.Append(fontSizeComplexScript73);

            paragraphProperties48.Append(justification26);
            paragraphProperties48.Append(paragraphMarkRunProperties48);
            ProofError proofError19 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run40 = new Run();

            RunProperties runProperties40 = new RunProperties();
            RunFonts runFonts88 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize75 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript74 = new FontSizeComplexScript() { Val = "20" };

            runProperties40.Append(runFonts88);
            runProperties40.Append(fontSize75);
            runProperties40.Append(fontSizeComplexScript74);
            Text text40 = new Text();
            text40.Text = labSheetA1Sheet.Positive35;

            run40.Append(runProperties40);
            run40.Append(text40);
            ProofError proofError20 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph48.Append(paragraphProperties48);
            paragraph48.Append(proofError19);
            paragraph48.Append(run40);
            paragraph48.Append(proofError20);

            tableCell43.Append(tableCellProperties43);
            tableCell43.Append(paragraph48);

            TableCell tableCell44 = new TableCell();

            TableCellProperties tableCellProperties44 = new TableCellProperties();
            TableCellWidth tableCellWidth44 = new TableCellWidth() { Width = "1087", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders44 = new TableCellBorders();
            LeftBorder leftBorder45 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder45 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders44.Append(leftBorder45);
            tableCellBorders44.Append(rightBorder45);
            Shading shading44 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment44 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties44.Append(tableCellWidth44);
            tableCellProperties44.Append(tableCellBorders44);
            tableCellProperties44.Append(shading44);
            tableCellProperties44.Append(tableCellVerticalAlignment44);

            Paragraph paragraph49 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties49 = new ParagraphProperties();
            Justification justification27 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties49 = new ParagraphMarkRunProperties();
            RunFonts runFonts89 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic9 = new Italic();
            FontSize fontSize76 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript75 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties49.Append(runFonts89);
            paragraphMarkRunProperties49.Append(italic9);
            paragraphMarkRunProperties49.Append(fontSize76);
            paragraphMarkRunProperties49.Append(fontSizeComplexScript75);

            paragraphProperties49.Append(justification27);
            paragraphProperties49.Append(paragraphMarkRunProperties49);
            ProofError proofError21 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run41 = new Run();

            RunProperties runProperties41 = new RunProperties();
            RunFonts runFonts90 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize77 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript76 = new FontSizeComplexScript() { Val = "20" };

            runProperties41.Append(runFonts90);
            runProperties41.Append(fontSize77);
            runProperties41.Append(fontSizeComplexScript76);
            Text text41 = new Text();
            text41.Text = labSheetA1Sheet.NonTarget35;

            run41.Append(runProperties41);
            run41.Append(text41);
            ProofError proofError22 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph49.Append(paragraphProperties49);
            paragraph49.Append(proofError21);
            paragraph49.Append(run41);
            paragraph49.Append(proofError22);

            tableCell44.Append(tableCellProperties44);
            tableCell44.Append(paragraph49);

            TableCell tableCell45 = new TableCell();

            TableCellProperties tableCellProperties45 = new TableCellProperties();
            TableCellWidth tableCellWidth45 = new TableCellWidth() { Width = "1007", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders45 = new TableCellBorders();
            LeftBorder leftBorder46 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder46 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders45.Append(leftBorder46);
            tableCellBorders45.Append(rightBorder46);
            Shading shading45 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment45 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties45.Append(tableCellWidth45);
            tableCellProperties45.Append(tableCellBorders45);
            tableCellProperties45.Append(shading45);
            tableCellProperties45.Append(tableCellVerticalAlignment45);

            Paragraph paragraph50 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties50 = new ParagraphProperties();
            Justification justification28 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties50 = new ParagraphMarkRunProperties();
            RunFonts runFonts91 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic10 = new Italic();
            FontSize fontSize78 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript77 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties50.Append(runFonts91);
            paragraphMarkRunProperties50.Append(italic10);
            paragraphMarkRunProperties50.Append(fontSize78);
            paragraphMarkRunProperties50.Append(fontSizeComplexScript77);

            paragraphProperties50.Append(justification28);
            paragraphProperties50.Append(paragraphMarkRunProperties50);
            ProofError proofError23 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run42 = new Run();

            RunProperties runProperties42 = new RunProperties();
            RunFonts runFonts92 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize79 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript78 = new FontSizeComplexScript() { Val = "20" };

            runProperties42.Append(runFonts92);
            runProperties42.Append(fontSize79);
            runProperties42.Append(fontSizeComplexScript78);
            Text text42 = new Text();
            text42.Text = labSheetA1Sheet.Negative35;

            run42.Append(runProperties42);
            run42.Append(text42);
            ProofError proofError24 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph50.Append(paragraphProperties50);
            paragraph50.Append(proofError23);
            paragraph50.Append(run42);
            paragraph50.Append(proofError24);

            tableCell45.Append(tableCellProperties45);
            tableCell45.Append(paragraph50);

            TableCell tableCell46 = new TableCell();

            TableCellProperties tableCellProperties46 = new TableCellProperties();
            TableCellWidth tableCellWidth46 = new TableCellWidth() { Width = "906", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders46 = new TableCellBorders();
            TopBorder topBorder31 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder47 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder28 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder47 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders46.Append(topBorder31);
            tableCellBorders46.Append(leftBorder47);
            tableCellBorders46.Append(bottomBorder28);
            tableCellBorders46.Append(rightBorder47);
            Shading shading46 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment46 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties46.Append(tableCellWidth46);
            tableCellProperties46.Append(tableCellBorders46);
            tableCellProperties46.Append(shading46);
            tableCellProperties46.Append(tableCellVerticalAlignment46);

            Paragraph paragraph51 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties51 = new ParagraphProperties();
            Justification justification29 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties51 = new ParagraphMarkRunProperties();
            RunFonts runFonts93 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize80 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript79 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties51.Append(runFonts93);
            paragraphMarkRunProperties51.Append(fontSize80);
            paragraphMarkRunProperties51.Append(fontSizeComplexScript79);

            paragraphProperties51.Append(justification29);
            paragraphProperties51.Append(paragraphMarkRunProperties51);
            ProofError proofError25 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run43 = new Run();

            RunProperties runProperties43 = new RunProperties();
            RunFonts runFonts94 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize81 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript80 = new FontSizeComplexScript() { Val = "20" };

            runProperties43.Append(runFonts94);
            runProperties43.Append(fontSize81);
            runProperties43.Append(fontSizeComplexScript80);
            Text text43 = new Text();
            text43.Text = labSheetA1Sheet.Blank35;

            run43.Append(runProperties43);
            run43.Append(text43);
            ProofError proofError26 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph51.Append(paragraphProperties51);
            paragraph51.Append(proofError25);
            paragraph51.Append(run43);
            paragraph51.Append(proofError26);

            tableCell46.Append(tableCellProperties46);
            tableCell46.Append(paragraph51);

            TableCell tableCell47 = new TableCell();

            TableCellProperties tableCellProperties47 = new TableCellProperties();
            TableCellWidth tableCellWidth47 = new TableCellWidth() { Width = "906", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders47 = new TableCellBorders();
            TopBorder topBorder32 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder48 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder29 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder48 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders47.Append(topBorder32);
            tableCellBorders47.Append(leftBorder48);
            tableCellBorders47.Append(bottomBorder29);
            tableCellBorders47.Append(rightBorder48);
            Shading shading47 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment47 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties47.Append(tableCellWidth47);
            tableCellProperties47.Append(tableCellBorders47);
            tableCellProperties47.Append(shading47);
            tableCellProperties47.Append(tableCellVerticalAlignment47);

            Paragraph paragraph52 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties52 = new ParagraphProperties();
            Justification justification30 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties52 = new ParagraphMarkRunProperties();
            RunFonts runFonts95 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize82 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript81 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties52.Append(runFonts95);
            paragraphMarkRunProperties52.Append(fontSize82);
            paragraphMarkRunProperties52.Append(fontSizeComplexScript81);

            paragraphProperties52.Append(justification30);
            paragraphProperties52.Append(paragraphMarkRunProperties52);

            Run run44 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties44 = new RunProperties();
            RunFonts runFonts96 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize83 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript82 = new FontSizeComplexScript() { Val = "20" };

            runProperties44.Append(runFonts96);
            runProperties44.Append(fontSize83);
            runProperties44.Append(fontSizeComplexScript82);
            Text text44 = new Text();
            text44.Text = "1X";

            run44.Append(runProperties44);
            run44.Append(text44);

            paragraph52.Append(paragraphProperties52);
            paragraph52.Append(run44);

            tableCell47.Append(tableCellProperties47);
            tableCell47.Append(paragraph52);

            TableCell tableCell48 = new TableCell();

            TableCellProperties tableCellProperties48 = new TableCellProperties();
            TableCellWidth tableCellWidth48 = new TableCellWidth() { Width = "1631", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders48 = new TableCellBorders();
            TopBorder topBorder33 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder49 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder30 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder49 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders48.Append(topBorder33);
            tableCellBorders48.Append(leftBorder49);
            tableCellBorders48.Append(bottomBorder30);
            tableCellBorders48.Append(rightBorder49);
            Shading shading48 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment48 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties48.Append(tableCellWidth48);
            tableCellProperties48.Append(tableCellBorders48);
            tableCellProperties48.Append(shading48);
            tableCellProperties48.Append(tableCellVerticalAlignment48);

            Paragraph paragraph53 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties53 = new ParagraphProperties();
            Justification justification31 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties53 = new ParagraphMarkRunProperties();
            RunFonts runFonts97 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize84 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript83 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties53.Append(runFonts97);
            paragraphMarkRunProperties53.Append(fontSize84);
            paragraphMarkRunProperties53.Append(fontSizeComplexScript83);

            paragraphProperties53.Append(justification31);
            paragraphProperties53.Append(paragraphMarkRunProperties53);
            ProofError proofError27 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run45 = new Run();

            RunProperties runProperties45 = new RunProperties();
            RunFonts runFonts98 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize85 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript84 = new FontSizeComplexScript() { Val = "20" };

            runProperties45.Append(runFonts98);
            runProperties45.Append(fontSize85);
            runProperties45.Append(fontSizeComplexScript84);
            Text text45 = new Text();
            text45.Text = labSheetA1Sheet.Lot35;

            run45.Append(runProperties45);
            run45.Append(text45);
            ProofError proofError28 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph53.Append(paragraphProperties53);
            paragraph53.Append(proofError27);
            paragraph53.Append(run45);
            paragraph53.Append(proofError28);

            tableCell48.Append(tableCellProperties48);
            tableCell48.Append(paragraph53);

            tableRow4.Append(tableCell34);
            tableRow4.Append(tableCell35);
            tableRow4.Append(tableCell36);
            tableRow4.Append(tableCell37);
            tableRow4.Append(tableCell38);
            tableRow4.Append(tableCell39);
            tableRow4.Append(tableCell40);
            tableRow4.Append(tableCell41);
            tableRow4.Append(tableCell42);
            tableRow4.Append(tableCell43);
            tableRow4.Append(tableCell44);
            tableRow4.Append(tableCell45);
            tableRow4.Append(tableCell46);
            tableRow4.Append(tableCell47);
            tableRow4.Append(tableCell48);

            TableRow tableRow5 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00397699", RsidTableRowProperties = "0032114E" };

            TableRowProperties tableRowProperties3 = new TableRowProperties();
            TableRowHeight tableRowHeight3 = new TableRowHeight() { Val = (UInt32Value)224U };

            tableRowProperties3.Append(tableRowHeight3);

            TableCell tableCell49 = new TableCell();

            TableCellProperties tableCellProperties49 = new TableCellProperties();
            TableCellWidth tableCellWidth49 = new TableCellWidth() { Width = "1031", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders49 = new TableCellBorders();
            TopBorder topBorder34 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder50 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder31 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder50 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders49.Append(topBorder34);
            tableCellBorders49.Append(leftBorder50);
            tableCellBorders49.Append(bottomBorder31);
            tableCellBorders49.Append(rightBorder50);
            Shading shading49 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment49 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties49.Append(tableCellWidth49);
            tableCellProperties49.Append(tableCellBorders49);
            tableCellProperties49.Append(shading49);
            tableCellProperties49.Append(tableCellVerticalAlignment49);

            Paragraph paragraph54 = new Paragraph() { RsidParagraphMarkRevision = "00030688", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties54 = new ParagraphProperties();
            Justification justification32 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties54 = new ParagraphMarkRunProperties();
            RunFonts runFonts99 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize86 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript85 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties54.Append(runFonts99);
            paragraphMarkRunProperties54.Append(fontSize86);
            paragraphMarkRunProperties54.Append(fontSizeComplexScript85);

            paragraphProperties54.Append(justification32);
            paragraphProperties54.Append(paragraphMarkRunProperties54);

            Run run46 = new Run() { RsidRunProperties = "00030688" };

            RunProperties runProperties46 = new RunProperties();
            RunFonts runFonts100 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize87 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript86 = new FontSizeComplexScript() { Val = "20" };

            runProperties46.Append(runFonts100);
            runProperties46.Append(fontSize87);
            runProperties46.Append(fontSizeComplexScript86);
            Text text46 = new Text();
            text46.Text = "Sample Crew";

            run46.Append(runProperties46);
            run46.Append(text46);

            paragraph54.Append(paragraphProperties54);
            paragraph54.Append(run46);

            tableCell49.Append(tableCellProperties49);
            tableCell49.Append(paragraph54);

            TableCell tableCell50 = new TableCell();

            TableCellProperties tableCellProperties50 = new TableCellProperties();
            TableCellWidth tableCellWidth50 = new TableCellWidth() { Width = "992", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders50 = new TableCellBorders();
            TopBorder topBorder35 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder51 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder32 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder51 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders50.Append(topBorder35);
            tableCellBorders50.Append(leftBorder51);
            tableCellBorders50.Append(bottomBorder32);
            tableCellBorders50.Append(rightBorder51);
            Shading shading50 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment50 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties50.Append(tableCellWidth50);
            tableCellProperties50.Append(tableCellBorders50);
            tableCellProperties50.Append(shading50);
            tableCellProperties50.Append(tableCellVerticalAlignment50);

            Paragraph paragraph55 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties55 = new ParagraphProperties();
            Justification justification33 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties55 = new ParagraphMarkRunProperties();
            RunFonts runFonts101 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize88 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript87 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties55.Append(runFonts101);
            paragraphMarkRunProperties55.Append(fontSize88);
            paragraphMarkRunProperties55.Append(fontSizeComplexScript87);

            paragraphProperties55.Append(justification33);
            paragraphProperties55.Append(paragraphMarkRunProperties55);
            ProofError proofError29 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run47 = new Run();

            RunProperties runProperties47 = new RunProperties();
            RunFonts runFonts102 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize89 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript88 = new FontSizeComplexScript() { Val = "20" };

            runProperties47.Append(runFonts102);
            runProperties47.Append(fontSize89);
            runProperties47.Append(fontSizeComplexScript88);
            Text text47 = new Text();
            text47.Text = labSheetA1Sheet.SampleCrewInitials;

            run47.Append(runProperties47);
            run47.Append(text47);
            ProofError proofError30 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph55.Append(paragraphProperties55);
            paragraph55.Append(proofError29);
            paragraph55.Append(run47);
            paragraph55.Append(proofError30);

            tableCell50.Append(tableCellProperties50);
            tableCell50.Append(paragraph55);

            TableCell tableCell51 = new TableCell();

            TableCellProperties tableCellProperties51 = new TableCellProperties();
            TableCellWidth tableCellWidth51 = new TableCellWidth() { Width = "284", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders51 = new TableCellBorders();
            TopBorder topBorder36 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder52 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder33 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder52 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders51.Append(topBorder36);
            tableCellBorders51.Append(leftBorder52);
            tableCellBorders51.Append(bottomBorder33);
            tableCellBorders51.Append(rightBorder52);
            Shading shading51 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment51 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties51.Append(tableCellWidth51);
            tableCellProperties51.Append(tableCellBorders51);
            tableCellProperties51.Append(shading51);
            tableCellProperties51.Append(tableCellVerticalAlignment51);

            Paragraph paragraph56 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties56 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties56 = new ParagraphMarkRunProperties();
            RunFonts runFonts103 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize90 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript89 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties56.Append(runFonts103);
            paragraphMarkRunProperties56.Append(fontSize90);
            paragraphMarkRunProperties56.Append(fontSizeComplexScript89);

            paragraphProperties56.Append(paragraphMarkRunProperties56);

            paragraph56.Append(paragraphProperties56);

            tableCell51.Append(tableCellProperties51);
            tableCell51.Append(paragraph56);

            TableCell tableCell52 = new TableCell();

            TableCellProperties tableCellProperties52 = new TableCellProperties();
            TableCellWidth tableCellWidth52 = new TableCellWidth() { Width = "1417", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders52 = new TableCellBorders();
            TopBorder topBorder37 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder53 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder34 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder53 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders52.Append(topBorder37);
            tableCellBorders52.Append(leftBorder53);
            tableCellBorders52.Append(bottomBorder34);
            tableCellBorders52.Append(rightBorder53);
            Shading shading52 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment52 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties52.Append(tableCellWidth52);
            tableCellProperties52.Append(tableCellBorders52);
            tableCellProperties52.Append(shading52);
            tableCellProperties52.Append(tableCellVerticalAlignment52);

            Paragraph paragraph57 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties57 = new ParagraphProperties();
            Justification justification34 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties57 = new ParagraphMarkRunProperties();
            RunFonts runFonts104 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize91 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript90 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties57.Append(runFonts104);
            paragraphMarkRunProperties57.Append(fontSize91);
            paragraphMarkRunProperties57.Append(fontSizeComplexScript90);

            paragraphProperties57.Append(justification34);
            paragraphProperties57.Append(paragraphMarkRunProperties57);

            Run run48 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties48 = new RunProperties();
            RunFonts runFonts105 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize92 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript91 = new FontSizeComplexScript() { Val = "20" };

            runProperties48.Append(runFonts105);
            runProperties48.Append(fontSize92);
            runProperties48.Append(fontSizeComplexScript91);
            Text text48 = new Text();
            text48.Text = "Water";

            run48.Append(runProperties48);
            run48.Append(text48);

            Run run49 = new Run() { RsidRunProperties = "00ED7FB2", RsidRunAddition = "00C47BCB" };

            RunProperties runProperties49 = new RunProperties();
            RunFonts runFonts106 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize93 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript92 = new FontSizeComplexScript() { Val = "20" };

            runProperties49.Append(runFonts106);
            runProperties49.Append(fontSize93);
            runProperties49.Append(fontSizeComplexScript92);
            Text text49 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text49.Text = " ";

            run49.Append(runProperties49);
            run49.Append(text49);

            Run run50 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties50 = new RunProperties();
            RunFonts runFonts107 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize94 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript93 = new FontSizeComplexScript() { Val = "20" };

            runProperties50.Append(runFonts107);
            runProperties50.Append(fontSize94);
            runProperties50.Append(fontSizeComplexScript93);
            Text text50 = new Text();
            text50.Text = "bath #";

            run50.Append(runProperties50);
            run50.Append(text50);

            paragraph57.Append(paragraphProperties57);
            paragraph57.Append(run48);
            paragraph57.Append(run49);
            paragraph57.Append(run50);

            tableCell52.Append(tableCellProperties52);
            tableCell52.Append(paragraph57);

            TableCell tableCell53 = new TableCell();

            TableCellProperties tableCellProperties53 = new TableCellProperties();
            TableCellWidth tableCellWidth53 = new TableCellWidth() { Width = "1985", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders53 = new TableCellBorders();
            TopBorder topBorder38 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder54 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder35 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder54 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders53.Append(topBorder38);
            tableCellBorders53.Append(leftBorder54);
            tableCellBorders53.Append(bottomBorder35);
            tableCellBorders53.Append(rightBorder54);
            Shading shading53 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment53 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties53.Append(tableCellWidth53);
            tableCellProperties53.Append(tableCellBorders53);
            tableCellProperties53.Append(shading53);
            tableCellProperties53.Append(tableCellVerticalAlignment53);

            Paragraph paragraph58 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties58 = new ParagraphProperties();
            Justification justification35 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties58 = new ParagraphMarkRunProperties();
            RunFonts runFonts108 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize95 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript94 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties58.Append(runFonts108);
            paragraphMarkRunProperties58.Append(fontSize95);
            paragraphMarkRunProperties58.Append(fontSizeComplexScript94);

            paragraphProperties58.Append(justification35);
            paragraphProperties58.Append(paragraphMarkRunProperties58);
            ProofError proofError31 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run51 = new Run();

            RunProperties runProperties51 = new RunProperties();
            RunFonts runFonts109 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize96 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript95 = new FontSizeComplexScript() { Val = "20" };

            runProperties51.Append(runFonts109);
            runProperties51.Append(fontSize96);
            runProperties51.Append(fontSizeComplexScript95);
            Text text51 = new Text();
            text51.Text = labSheetA1Sheet.WaterBath1 +
                (labSheetA1Sheet.WaterBathCount > 1 ? "/" + labSheetA1Sheet.WaterBath2 : "") +
                (labSheetA1Sheet.WaterBathCount > 2 ? "/" + labSheetA1Sheet.WaterBath3 : "");

            run51.Append(runProperties51);
            run51.Append(text51);
            ProofError proofError32 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph58.Append(paragraphProperties58);
            paragraph58.Append(proofError31);
            paragraph58.Append(run51);
            paragraph58.Append(proofError32);

            tableCell53.Append(tableCellProperties53);
            tableCell53.Append(paragraph58);

            TableCell tableCell54 = new TableCell();

            TableCellProperties tableCellProperties54 = new TableCellProperties();
            TableCellWidth tableCellWidth54 = new TableCellWidth() { Width = "283", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders54 = new TableCellBorders();
            TopBorder topBorder39 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder55 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder36 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder55 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders54.Append(topBorder39);
            tableCellBorders54.Append(leftBorder55);
            tableCellBorders54.Append(bottomBorder36);
            tableCellBorders54.Append(rightBorder55);
            Shading shading54 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment54 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties54.Append(tableCellWidth54);
            tableCellProperties54.Append(tableCellBorders54);
            tableCellProperties54.Append(shading54);
            tableCellProperties54.Append(tableCellVerticalAlignment54);

            Paragraph paragraph59 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties59 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties59 = new ParagraphMarkRunProperties();
            RunFonts runFonts110 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color6 = new Color() { Val = "DDDDDD" };
            FontSize fontSize97 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript96 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties59.Append(runFonts110);
            paragraphMarkRunProperties59.Append(color6);
            paragraphMarkRunProperties59.Append(fontSize97);
            paragraphMarkRunProperties59.Append(fontSizeComplexScript96);

            paragraphProperties59.Append(paragraphMarkRunProperties59);

            paragraph59.Append(paragraphProperties59);

            tableCell54.Append(tableCellProperties54);
            tableCell54.Append(paragraph59);

            TableCell tableCell55 = new TableCell();

            TableCellProperties tableCellProperties55 = new TableCellProperties();
            TableCellWidth tableCellWidth55 = new TableCellWidth() { Width = "1416", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders55 = new TableCellBorders();
            TopBorder topBorder40 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder56 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder37 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder56 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders55.Append(topBorder40);
            tableCellBorders55.Append(leftBorder56);
            tableCellBorders55.Append(bottomBorder37);
            tableCellBorders55.Append(rightBorder56);
            Shading shading55 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment55 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties55.Append(tableCellWidth55);
            tableCellProperties55.Append(tableCellBorders55);
            tableCellProperties55.Append(shading55);
            tableCellProperties55.Append(tableCellVerticalAlignment55);

            Paragraph paragraph60 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties60 = new ParagraphProperties();
            Justification justification36 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties60 = new ParagraphMarkRunProperties();
            RunFonts runFonts111 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color7 = new Color() { Val = "C0C0C0" };
            FontSize fontSize98 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript97 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties60.Append(runFonts111);
            paragraphMarkRunProperties60.Append(color7);
            paragraphMarkRunProperties60.Append(fontSize98);
            paragraphMarkRunProperties60.Append(fontSizeComplexScript97);

            paragraphProperties60.Append(justification36);
            paragraphProperties60.Append(paragraphMarkRunProperties60);
            ProofError proofError33 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run52 = new Run();

            RunProperties runProperties52 = new RunProperties();
            RunFonts runFonts112 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize99 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript98 = new FontSizeComplexScript() { Val = "20" };

            runProperties52.Append(runFonts112);
            runProperties52.Append(fontSize99);
            runProperties52.Append(fontSizeComplexScript98);
            Text text52 = new Text();
            text52.Text = labSheetA1Sheet.TCLab1 + (labSheetA1Sheet.TCLab2 != null && labSheetA1Sheet.TCLab2.Length > 0 ? " / " + labSheetA1Sheet.TCLab2 : "");

            run52.Append(runProperties52);
            run52.Append(text52);
            ProofError proofError34 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph60.Append(paragraphProperties60);
            paragraph60.Append(proofError33);
            paragraph60.Append(run52);
            paragraph60.Append(proofError34);

            tableCell55.Append(tableCellProperties55);
            tableCell55.Append(paragraph60);

            TableCell tableCell56 = new TableCell();

            TableCellProperties tableCellProperties56 = new TableCellProperties();
            TableCellWidth tableCellWidth56 = new TableCellWidth() { Width = "238", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders56 = new TableCellBorders();
            TopBorder topBorder41 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder57 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder38 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder57 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders56.Append(topBorder41);
            tableCellBorders56.Append(leftBorder57);
            tableCellBorders56.Append(bottomBorder38);
            tableCellBorders56.Append(rightBorder57);
            Shading shading56 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment56 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties56.Append(tableCellWidth56);
            tableCellProperties56.Append(tableCellBorders56);
            tableCellProperties56.Append(shading56);
            tableCellProperties56.Append(tableCellVerticalAlignment56);

            Paragraph paragraph61 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties61 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties61 = new ParagraphMarkRunProperties();
            RunFonts runFonts113 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic11 = new Italic();
            FontSize fontSize100 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript99 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties61.Append(runFonts113);
            paragraphMarkRunProperties61.Append(italic11);
            paragraphMarkRunProperties61.Append(fontSize100);
            paragraphMarkRunProperties61.Append(fontSizeComplexScript99);

            paragraphProperties61.Append(paragraphMarkRunProperties61);

            paragraph61.Append(paragraphProperties61);

            tableCell56.Append(tableCellProperties56);
            tableCell56.Append(paragraph61);

            TableCell tableCell57 = new TableCell();

            TableCellProperties tableCellProperties57 = new TableCellProperties();
            TableCellWidth tableCellWidth57 = new TableCellWidth() { Width = "1031", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders57 = new TableCellBorders();
            TopBorder topBorder42 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder58 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder39 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder58 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders57.Append(topBorder42);
            tableCellBorders57.Append(leftBorder58);
            tableCellBorders57.Append(bottomBorder39);
            tableCellBorders57.Append(rightBorder58);
            Shading shading57 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment57 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties57.Append(tableCellWidth57);
            tableCellProperties57.Append(tableCellBorders57);
            tableCellProperties57.Append(shading57);
            tableCellProperties57.Append(tableCellVerticalAlignment57);

            Paragraph paragraph62 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0055454A", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties62 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties62 = new ParagraphMarkRunProperties();
            RunFonts runFonts114 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize101 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript100 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties62.Append(runFonts114);
            paragraphMarkRunProperties62.Append(fontSize101);
            paragraphMarkRunProperties62.Append(fontSizeComplexScript100);

            paragraphProperties62.Append(paragraphMarkRunProperties62);

            Run run53 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties53 = new RunProperties();
            RunFonts runFonts115 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize102 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript101 = new FontSizeComplexScript() { Val = "20" };

            runProperties53.Append(runFonts115);
            runProperties53.Append(fontSize102);
            runProperties53.Append(fontSizeComplexScript101);
            Text text53 = new Text();
            text53.Text = "44.5";

            run53.Append(runProperties53);
            run53.Append(text53);

            Run run54 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties54 = new RunProperties();
            RunFonts runFonts116 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize103 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript102 = new FontSizeComplexScript() { Val = "20" };
            VerticalTextAlignment verticalTextAlignment4 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties54.Append(runFonts116);
            runProperties54.Append(fontSize103);
            runProperties54.Append(fontSizeComplexScript102);
            runProperties54.Append(verticalTextAlignment4);
            Text text54 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text54.Text = " ";

            run54.Append(runProperties54);
            run54.Append(text54);
            ProofError proofError35 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run55 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties55 = new RunProperties();
            RunFonts runFonts117 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize104 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript103 = new FontSizeComplexScript() { Val = "20" };
            VerticalTextAlignment verticalTextAlignment5 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties55.Append(runFonts117);
            runProperties55.Append(fontSize104);
            runProperties55.Append(fontSizeComplexScript103);
            runProperties55.Append(verticalTextAlignment5);
            Text text55 = new Text();
            text55.Text = "o";

            run55.Append(runProperties55);
            run55.Append(text55);

            Run run56 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties56 = new RunProperties();
            RunFonts runFonts118 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize105 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript104 = new FontSizeComplexScript() { Val = "20" };

            runProperties56.Append(runFonts118);
            runProperties56.Append(fontSize105);
            runProperties56.Append(fontSizeComplexScript104);
            Text text56 = new Text();
            text56.Text = "C";

            run56.Append(runProperties56);
            run56.Append(text56);
            ProofError proofError36 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph62.Append(paragraphProperties62);
            paragraph62.Append(run53);
            paragraph62.Append(run54);
            paragraph62.Append(proofError35);
            paragraph62.Append(run55);
            paragraph62.Append(run56);
            paragraph62.Append(proofError36);

            tableCell57.Append(tableCellProperties57);
            tableCell57.Append(paragraph62);

            TableCell tableCell58 = new TableCell();

            TableCellProperties tableCellProperties58 = new TableCellProperties();
            TableCellWidth tableCellWidth58 = new TableCellWidth() { Width = "906", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders58 = new TableCellBorders();
            LeftBorder leftBorder59 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder40 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder59 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders58.Append(leftBorder59);
            tableCellBorders58.Append(bottomBorder40);
            tableCellBorders58.Append(rightBorder59);
            Shading shading58 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment58 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties58.Append(tableCellWidth58);
            tableCellProperties58.Append(tableCellBorders58);
            tableCellProperties58.Append(shading58);
            tableCellProperties58.Append(tableCellVerticalAlignment58);

            Paragraph paragraph63 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties63 = new ParagraphProperties();
            Justification justification37 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties63 = new ParagraphMarkRunProperties();
            RunFonts runFonts119 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic12 = new Italic();
            FontSize fontSize106 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript105 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties63.Append(runFonts119);
            paragraphMarkRunProperties63.Append(italic12);
            paragraphMarkRunProperties63.Append(fontSize106);
            paragraphMarkRunProperties63.Append(fontSizeComplexScript105);

            paragraphProperties63.Append(justification37);
            paragraphProperties63.Append(paragraphMarkRunProperties63);
            ProofError proofError37 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run57 = new Run();

            RunProperties runProperties57 = new RunProperties();
            RunFonts runFonts120 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize107 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript106 = new FontSizeComplexScript() { Val = "20" };

            runProperties57.Append(runFonts120);
            runProperties57.Append(fontSize107);
            runProperties57.Append(fontSizeComplexScript106);
            Text text57 = new Text();
            text57.Text = labSheetA1Sheet.Bath1Positive44_5 +
                (labSheetA1Sheet.WaterBathCount > 1 ? "/" + labSheetA1Sheet.Bath2Positive44_5 : "") +
                (labSheetA1Sheet.WaterBathCount > 2 ? "/" + labSheetA1Sheet.Bath3Positive44_5 : "");

            run57.Append(runProperties57);
            run57.Append(text57);
            ProofError proofError38 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph63.Append(paragraphProperties63);
            paragraph63.Append(proofError37);
            paragraph63.Append(run57);
            paragraph63.Append(proofError38);

            tableCell58.Append(tableCellProperties58);
            tableCell58.Append(paragraph63);

            TableCell tableCell59 = new TableCell();

            TableCellProperties tableCellProperties59 = new TableCellProperties();
            TableCellWidth tableCellWidth59 = new TableCellWidth() { Width = "1087", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders59 = new TableCellBorders();
            LeftBorder leftBorder60 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder41 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder60 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders59.Append(leftBorder60);
            tableCellBorders59.Append(bottomBorder41);
            tableCellBorders59.Append(rightBorder60);
            Shading shading59 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment59 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties59.Append(tableCellWidth59);
            tableCellProperties59.Append(tableCellBorders59);
            tableCellProperties59.Append(shading59);
            tableCellProperties59.Append(tableCellVerticalAlignment59);

            Paragraph paragraph64 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties64 = new ParagraphProperties();
            Justification justification38 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties64 = new ParagraphMarkRunProperties();
            RunFonts runFonts121 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic13 = new Italic();
            FontSize fontSize108 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript107 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties64.Append(runFonts121);
            paragraphMarkRunProperties64.Append(italic13);
            paragraphMarkRunProperties64.Append(fontSize108);
            paragraphMarkRunProperties64.Append(fontSizeComplexScript107);

            paragraphProperties64.Append(justification38);
            paragraphProperties64.Append(paragraphMarkRunProperties64);
            ProofError proofError39 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run58 = new Run();

            RunProperties runProperties58 = new RunProperties();
            RunFonts runFonts122 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize109 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript108 = new FontSizeComplexScript() { Val = "20" };

            runProperties58.Append(runFonts122);
            runProperties58.Append(fontSize109);
            runProperties58.Append(fontSizeComplexScript108);
            Text text58 = new Text();
            text58.Text = labSheetA1Sheet.Bath1NonTarget44_5 +
                (labSheetA1Sheet.WaterBathCount > 1 ? "/" + labSheetA1Sheet.Bath2NonTarget44_5 : "") +
                (labSheetA1Sheet.WaterBathCount > 2 ? "/" + labSheetA1Sheet.Bath3NonTarget44_5 : ""); ;

            run58.Append(runProperties58);
            run58.Append(text58);
            ProofError proofError40 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph64.Append(paragraphProperties64);
            paragraph64.Append(proofError39);
            paragraph64.Append(run58);
            paragraph64.Append(proofError40);

            tableCell59.Append(tableCellProperties59);
            tableCell59.Append(paragraph64);

            TableCell tableCell60 = new TableCell();

            TableCellProperties tableCellProperties60 = new TableCellProperties();
            TableCellWidth tableCellWidth60 = new TableCellWidth() { Width = "1007", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders60 = new TableCellBorders();
            LeftBorder leftBorder61 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder42 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder61 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders60.Append(leftBorder61);
            tableCellBorders60.Append(bottomBorder42);
            tableCellBorders60.Append(rightBorder61);
            Shading shading60 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment60 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties60.Append(tableCellWidth60);
            tableCellProperties60.Append(tableCellBorders60);
            tableCellProperties60.Append(shading60);
            tableCellProperties60.Append(tableCellVerticalAlignment60);

            Paragraph paragraph65 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "0072154A", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties65 = new ParagraphProperties();
            Justification justification39 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties65 = new ParagraphMarkRunProperties();
            RunFonts runFonts123 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic14 = new Italic();
            FontSize fontSize110 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript109 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties65.Append(runFonts123);
            paragraphMarkRunProperties65.Append(italic14);
            paragraphMarkRunProperties65.Append(fontSize110);
            paragraphMarkRunProperties65.Append(fontSizeComplexScript109);

            paragraphProperties65.Append(justification39);
            paragraphProperties65.Append(paragraphMarkRunProperties65);
            ProofError proofError41 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run59 = new Run();

            RunProperties runProperties59 = new RunProperties();
            RunFonts runFonts124 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize111 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript110 = new FontSizeComplexScript() { Val = "20" };

            runProperties59.Append(runFonts124);
            runProperties59.Append(fontSize111);
            runProperties59.Append(fontSizeComplexScript110);
            Text text59 = new Text();
            text59.Text = labSheetA1Sheet.Bath1Negative44_5 +
                (labSheetA1Sheet.WaterBathCount > 1 ? "/" + labSheetA1Sheet.Bath2Negative44_5 : "") +
                (labSheetA1Sheet.WaterBathCount > 2 ? "/" + labSheetA1Sheet.Bath3Negative44_5 : ""); 

            run59.Append(runProperties59);
            run59.Append(text59);
            ProofError proofError42 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph65.Append(paragraphProperties65);
            paragraph65.Append(proofError41);
            paragraph65.Append(run59);
            paragraph65.Append(proofError42);

            tableCell60.Append(tableCellProperties60);
            tableCell60.Append(paragraph65);

            TableCell tableCell61 = new TableCell();

            TableCellProperties tableCellProperties61 = new TableCellProperties();
            TableCellWidth tableCellWidth61 = new TableCellWidth() { Width = "906", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders61 = new TableCellBorders();
            TopBorder topBorder43 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder62 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder43 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder62 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders61.Append(topBorder43);
            tableCellBorders61.Append(leftBorder62);
            tableCellBorders61.Append(bottomBorder43);
            tableCellBorders61.Append(rightBorder62);
            Shading shading61 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment61 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties61.Append(tableCellWidth61);
            tableCellProperties61.Append(tableCellBorders61);
            tableCellProperties61.Append(shading61);
            tableCellProperties61.Append(tableCellVerticalAlignment61);

            Paragraph paragraph66 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties66 = new ParagraphProperties();
            Justification justification40 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties66 = new ParagraphMarkRunProperties();
            RunFonts runFonts125 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize112 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript111 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties66.Append(runFonts125);
            paragraphMarkRunProperties66.Append(fontSize112);
            paragraphMarkRunProperties66.Append(fontSizeComplexScript111);

            paragraphProperties66.Append(justification40);
            paragraphProperties66.Append(paragraphMarkRunProperties66);
            ProofError proofError43 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run60 = new Run();

            RunProperties runProperties60 = new RunProperties();
            RunFonts runFonts126 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize113 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript112 = new FontSizeComplexScript() { Val = "20" };

            runProperties60.Append(runFonts126);
            runProperties60.Append(fontSize113);
            runProperties60.Append(fontSizeComplexScript112);
            Text text60 = new Text();
            text60.Text = labSheetA1Sheet.Bath1Blank44_5 +
                (labSheetA1Sheet.WaterBathCount > 1 ? "/" + labSheetA1Sheet.Bath2Blank44_5 : "") +
                (labSheetA1Sheet.WaterBathCount > 2 ? "/" + labSheetA1Sheet.Bath3Blank44_5 : ""); ;

            run60.Append(runProperties60);
            run60.Append(text60);
            ProofError proofError44 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph66.Append(paragraphProperties66);
            paragraph66.Append(proofError43);
            paragraph66.Append(run60);
            paragraph66.Append(proofError44);

            tableCell61.Append(tableCellProperties61);
            tableCell61.Append(paragraph66);

            TableCell tableCell62 = new TableCell();

            TableCellProperties tableCellProperties62 = new TableCellProperties();
            TableCellWidth tableCellWidth62 = new TableCellWidth() { Width = "906", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders62 = new TableCellBorders();
            TopBorder topBorder44 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder63 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder44 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder63 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders62.Append(topBorder44);
            tableCellBorders62.Append(leftBorder63);
            tableCellBorders62.Append(bottomBorder44);
            tableCellBorders62.Append(rightBorder63);
            Shading shading62 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment62 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties62.Append(tableCellWidth62);
            tableCellProperties62.Append(tableCellBorders62);
            tableCellProperties62.Append(shading62);
            tableCellProperties62.Append(tableCellVerticalAlignment62);

            Paragraph paragraph67 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00397699" };

            ParagraphProperties paragraphProperties67 = new ParagraphProperties();
            Justification justification41 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties67 = new ParagraphMarkRunProperties();
            RunFonts runFonts127 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize114 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript113 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties67.Append(runFonts127);
            paragraphMarkRunProperties67.Append(fontSize114);
            paragraphMarkRunProperties67.Append(fontSizeComplexScript113);

            paragraphProperties67.Append(justification41);
            paragraphProperties67.Append(paragraphMarkRunProperties67);

            Run run61 = new Run() { RsidRunProperties = "00ED7FB2" };

            RunProperties runProperties61 = new RunProperties();
            RunFonts runFonts128 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize115 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript114 = new FontSizeComplexScript() { Val = "20" };

            runProperties61.Append(runFonts128);
            runProperties61.Append(fontSize115);
            runProperties61.Append(fontSizeComplexScript114);
            Text text61 = new Text();
            text61.Text = "2X";

            run61.Append(runProperties61);
            run61.Append(text61);

            paragraph67.Append(paragraphProperties67);
            paragraph67.Append(run61);

            tableCell62.Append(tableCellProperties62);
            tableCell62.Append(paragraph67);

            TableCell tableCell63 = new TableCell();

            TableCellProperties tableCellProperties63 = new TableCellProperties();
            TableCellWidth tableCellWidth63 = new TableCellWidth() { Width = "1631", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders63 = new TableCellBorders();
            TopBorder topBorder45 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder64 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder45 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder64 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders63.Append(topBorder45);
            tableCellBorders63.Append(leftBorder64);
            tableCellBorders63.Append(bottomBorder45);
            tableCellBorders63.Append(rightBorder64);
            Shading shading63 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment63 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties63.Append(tableCellWidth63);
            tableCellProperties63.Append(tableCellBorders63);
            tableCellProperties63.Append(shading63);
            tableCellProperties63.Append(tableCellVerticalAlignment63);

            Paragraph paragraph68 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00397699", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties68 = new ParagraphProperties();
            Justification justification42 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties68 = new ParagraphMarkRunProperties();
            RunFonts runFonts129 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize116 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript115 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties68.Append(runFonts129);
            paragraphMarkRunProperties68.Append(fontSize116);
            paragraphMarkRunProperties68.Append(fontSizeComplexScript115);

            paragraphProperties68.Append(justification42);
            paragraphProperties68.Append(paragraphMarkRunProperties68);

            Run run62 = new Run();

            RunProperties runProperties62 = new RunProperties();
            RunFonts runFonts130 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize117 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript116 = new FontSizeComplexScript() { Val = "20" };

            runProperties62.Append(runFonts130);
            runProperties62.Append(fontSize117);
            runProperties62.Append(fontSizeComplexScript116);
            Text text62 = new Text();
            text62.Text = labSheetA1Sheet.Lot44_5;

            run62.Append(runProperties62);
            run62.Append(text62);

            paragraph68.Append(paragraphProperties68);
            paragraph68.Append(run62);

            tableCell63.Append(tableCellProperties63);
            tableCell63.Append(paragraph68);

            tableRow5.Append(tableRowProperties3);
            tableRow5.Append(tableCell49);
            tableRow5.Append(tableCell50);
            tableRow5.Append(tableCell51);
            tableRow5.Append(tableCell52);
            tableRow5.Append(tableCell53);
            tableRow5.Append(tableCell54);
            tableRow5.Append(tableCell55);
            tableRow5.Append(tableCell56);
            tableRow5.Append(tableCell57);
            tableRow5.Append(tableCell58);
            tableRow5.Append(tableCell59);
            tableRow5.Append(tableCell60);
            tableRow5.Append(tableCell61);
            tableRow5.Append(tableCell62);
            tableRow5.Append(tableCell63);

            table2.Append(tableProperties2);
            table2.Append(tableGrid2);
            table2.Append(tableRow2);
            table2.Append(tableRow3);
            table2.Append(tableRow4);
            table2.Append(tableRow5);
        }
    }
}