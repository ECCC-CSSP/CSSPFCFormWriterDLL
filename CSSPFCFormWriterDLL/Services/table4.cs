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
        public void DoTable4(Table table4)
        {
            TableProperties tableProperties4 = new TableProperties();
            TableWidth tableWidth4 = new TableWidth() { Width = "14760", Type = TableWidthUnitValues.Dxa };
            TableIndentation tableIndentation3 = new TableIndentation() { Width = -72, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders3 = new TableBorders();
            TopBorder topBorder402 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder435 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder270 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder435 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder3 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder3 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders3.Append(topBorder402);
            tableBorders3.Append(leftBorder435);
            tableBorders3.Append(bottomBorder270);
            tableBorders3.Append(rightBorder435);
            tableBorders3.Append(insideHorizontalBorder3);
            tableBorders3.Append(insideVerticalBorder3);
            TableLayout tableLayout3 = new TableLayout() { Type = TableLayoutValues.Fixed };
            TableLook tableLook4 = new TableLook() { Val = "01E0" };

            tableProperties4.Append(tableWidth4);
            tableProperties4.Append(tableIndentation3);
            tableProperties4.Append(tableBorders3);
            tableProperties4.Append(tableLayout3);
            tableProperties4.Append(tableLook4);

            TableGrid tableGrid4 = new TableGrid();
            GridColumn gridColumn41 = new GridColumn() { Width = "8224" };
            GridColumn gridColumn42 = new GridColumn() { Width = "236" };
            GridColumn gridColumn43 = new GridColumn() { Width = "2520" };
            GridColumn gridColumn44 = new GridColumn() { Width = "3780" };

            tableGrid4.Append(gridColumn41);
            tableGrid4.Append(gridColumn42);
            tableGrid4.Append(gridColumn43);
            tableGrid4.Append(gridColumn44);

            TableRow tableRow27 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "00A34788", RsidTableRowProperties = "00ED7FB2" };

            TableRowProperties tableRowProperties24 = new TableRowProperties();
            TableRowHeight tableRowHeight24 = new TableRowHeight() { Val = (UInt32Value)174U };

            tableRowProperties24.Append(tableRowHeight24);

            TableCell tableCell433 = new TableCell();

            TableCellProperties tableCellProperties433 = new TableCellProperties();
            TableCellWidth tableCellWidth433 = new TableCellWidth() { Width = "8224", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge50 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders433 = new TableCellBorders();
            TopBorder topBorder403 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder436 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder271 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder436 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders433.Append(topBorder403);
            tableCellBorders433.Append(leftBorder436);
            tableCellBorders433.Append(bottomBorder271);
            tableCellBorders433.Append(rightBorder436);
            Shading shading433 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment78 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties433.Append(tableCellWidth433);
            tableCellProperties433.Append(verticalMerge50);
            tableCellProperties433.Append(tableCellBorders433);
            tableCellProperties433.Append(shading433);
            tableCellProperties433.Append(tableCellVerticalAlignment78);

            Paragraph paragraph475 = new Paragraph() { RsidParagraphMarkRevision = "00322ED6", RsidParagraphAddition = "00A34788", RsidParagraphProperties = "00322ED6", RsidRunAdditionDefault = "00A34788" };

            ParagraphProperties paragraphProperties475 = new ParagraphProperties();
            Justification justification384 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties475 = new ParagraphMarkRunProperties();
            RunFonts runFonts605 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold7 = new Bold();
            FontSize fontSize264 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript262 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties475.Append(runFonts605);
            paragraphMarkRunProperties475.Append(bold7);
            paragraphMarkRunProperties475.Append(fontSize264);
            paragraphMarkRunProperties475.Append(fontSizeComplexScript262);

            paragraphProperties475.Append(justification384);
            paragraphProperties475.Append(paragraphMarkRunProperties475);

            Run run132 = new Run() { RsidRunProperties = "00322ED6" };

            RunProperties runProperties132 = new RunProperties();
            RunFonts runFonts606 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold8 = new Bold();
            FontSize fontSize265 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript263 = new FontSizeComplexScript() { Val = "22" };

            runProperties132.Append(runFonts606);
            runProperties132.Append(bold8);
            runProperties132.Append(fontSize265);
            runProperties132.Append(fontSizeComplexScript263);
            Text text132 = new Text();
            text132.Text = "Weather:";

            run132.Append(runProperties132);
            run132.Append(text132);

            Run run133 = new Run() { RsidRunProperties = "00322ED6", RsidRunAddition = "0012629E" };

            RunProperties runProperties133 = new RunProperties();
            RunFonts runFonts607 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold9 = new Bold();
            FontSize fontSize266 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript264 = new FontSizeComplexScript() { Val = "22" };

            runProperties133.Append(runFonts607);
            runProperties133.Append(bold9);
            runProperties133.Append(fontSize266);
            runProperties133.Append(fontSizeComplexScript264);
            Text text133 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text133.Text = "  ";

            run133.Append(runProperties133);
            run133.Append(text133);
            ProofError proofError55 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run134 = new Run() { RsidRunAddition = "00030688" };

            RunProperties runProperties134 = new RunProperties();
            RunFonts runFonts608 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            //Bold bold10 = new Bold();
            FontSize fontSize267 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript265 = new FontSizeComplexScript() { Val = "22" };

            runProperties134.Append(runFonts608);
            //runProperties134.Append(bold10);
            runProperties134.Append(fontSize267);
            runProperties134.Append(fontSizeComplexScript265);
            Text text134 = new Text();
            text134.Text = (labSheetA1Sheet.RunWeatherComment.Length > 140 ? labSheetA1Sheet.RunWeatherComment.Substring(0, 140) + " ..." : labSheetA1Sheet.RunWeatherComment);

            run134.Append(runProperties134);
            run134.Append(text134);
            ProofError proofError56 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph475.Append(paragraphProperties475);
            paragraph475.Append(run132);
            paragraph475.Append(run133);
            paragraph475.Append(proofError55);
            paragraph475.Append(run134);
            paragraph475.Append(proofError56);

            tableCell433.Append(tableCellProperties433);
            tableCell433.Append(paragraph475);

            TableCell tableCell434 = new TableCell();

            TableCellProperties tableCellProperties434 = new TableCellProperties();
            TableCellWidth tableCellWidth434 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders434 = new TableCellBorders();
            TopBorder topBorder404 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder437 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder272 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder437 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders434.Append(topBorder404);
            tableCellBorders434.Append(leftBorder437);
            tableCellBorders434.Append(bottomBorder272);
            tableCellBorders434.Append(rightBorder437);
            Shading shading434 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment79 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties434.Append(tableCellWidth434);
            tableCellProperties434.Append(tableCellBorders434);
            tableCellProperties434.Append(shading434);
            tableCellProperties434.Append(tableCellVerticalAlignment79);

            Paragraph paragraph476 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "00A34788", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00A34788" };

            ParagraphProperties paragraphProperties476 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties476 = new ParagraphMarkRunProperties();
            RunFonts runFonts609 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties476.Append(runFonts609);

            paragraphProperties476.Append(paragraphMarkRunProperties476);

            paragraph476.Append(paragraphProperties476);

            tableCell434.Append(tableCellProperties434);
            tableCell434.Append(paragraph476);

            TableCell tableCell435 = new TableCell();

            TableCellProperties tableCellProperties435 = new TableCellProperties();
            TableCellWidth tableCellWidth435 = new TableCellWidth() { Width = "2520", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders435 = new TableCellBorders();
            TopBorder topBorder405 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder438 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder273 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder438 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders435.Append(topBorder405);
            tableCellBorders435.Append(leftBorder438);
            tableCellBorders435.Append(bottomBorder273);
            tableCellBorders435.Append(rightBorder438);
            Shading shading435 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment80 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties435.Append(tableCellWidth435);
            tableCellProperties435.Append(tableCellBorders435);
            tableCellProperties435.Append(shading435);
            tableCellProperties435.Append(tableCellVerticalAlignment80);

            Paragraph paragraph477 = new Paragraph() { RsidParagraphMarkRevision = "00322ED6", RsidParagraphAddition = "00A34788", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties477 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties477 = new ParagraphMarkRunProperties();
            RunFonts runFonts610 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold11 = new Bold();
            FontSize fontSize268 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript266 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties477.Append(runFonts610);
            paragraphMarkRunProperties477.Append(bold11);
            paragraphMarkRunProperties477.Append(fontSize268);
            paragraphMarkRunProperties477.Append(fontSizeComplexScript266);

            paragraphProperties477.Append(paragraphMarkRunProperties477);

            Run run135 = new Run() { RsidRunProperties = "00322ED6" };

            RunProperties runProperties135 = new RunProperties();
            RunFonts runFonts611 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold12 = new Bold();
            FontSize fontSize269 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript267 = new FontSizeComplexScript() { Val = "22" };

            runProperties135.Append(runFonts611);
            runProperties135.Append(bold12);
            runProperties135.Append(fontSize269);
            runProperties135.Append(fontSizeComplexScript267);
            Text text135 = new Text();
            text135.Text = "Sample Bottle Lot #";

            run135.Append(runProperties135);
            run135.Append(text135);

            paragraph477.Append(paragraphProperties477);
            paragraph477.Append(run135);

            tableCell435.Append(tableCellProperties435);
            tableCell435.Append(paragraph477);

            TableCell tableCell436 = new TableCell();

            TableCellProperties tableCellProperties436 = new TableCellProperties();
            TableCellWidth tableCellWidth436 = new TableCellWidth() { Width = "3780", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders436 = new TableCellBorders();
            TopBorder topBorder406 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder439 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder274 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder439 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders436.Append(topBorder406);
            tableCellBorders436.Append(leftBorder439);
            tableCellBorders436.Append(bottomBorder274);
            tableCellBorders436.Append(rightBorder439);
            Shading shading436 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment81 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties436.Append(tableCellWidth436);
            tableCellProperties436.Append(tableCellBorders436);
            tableCellProperties436.Append(shading436);
            tableCellProperties436.Append(tableCellVerticalAlignment81);

            Paragraph paragraph478 = new Paragraph() { RsidParagraphMarkRevision = "00322ED6", RsidParagraphAddition = "00A34788", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties478 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties478 = new ParagraphMarkRunProperties();
            RunFonts runFonts612 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color36 = new Color() { Val = "C0C0C0" };
            FontSize fontSize270 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript268 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties478.Append(runFonts612);
            paragraphMarkRunProperties478.Append(color36);
            paragraphMarkRunProperties478.Append(fontSize270);
            paragraphMarkRunProperties478.Append(fontSizeComplexScript268);

            paragraphProperties478.Append(paragraphMarkRunProperties478);

            Run run136 = new Run();

            RunProperties runProperties136 = new RunProperties();
            RunFonts runFonts613 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize271 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript269 = new FontSizeComplexScript() { Val = "20" };

            runProperties136.Append(runFonts613);
            runProperties136.Append(fontSize271);
            runProperties136.Append(fontSizeComplexScript269);
            Text text136 = new Text();
            text136.Text = labSheetA1Sheet.SampleBottleLotNumber;

            run136.Append(runProperties136);
            run136.Append(text136);

            paragraph478.Append(paragraphProperties478);
            paragraph478.Append(run136);

            tableCell436.Append(tableCellProperties436);
            tableCell436.Append(paragraph478);

            tableRow27.Append(tableRowProperties24);
            tableRow27.Append(tableCell433);
            tableRow27.Append(tableCell434);
            tableRow27.Append(tableCell435);
            tableRow27.Append(tableCell436);

            TableRow tableRow28 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "009A6852", RsidTableRowProperties = "0032114E" };

            TableRowProperties tableRowProperties25 = new TableRowProperties();
            TableRowHeight tableRowHeight25 = new TableRowHeight() { Val = (UInt32Value)352U };

            tableRowProperties25.Append(tableRowHeight25);

            TableCell tableCell437 = new TableCell();

            TableCellProperties tableCellProperties437 = new TableCellProperties();
            TableCellWidth tableCellWidth437 = new TableCellWidth() { Width = "8224", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge51 = new VerticalMerge();

            TableCellBorders tableCellBorders437 = new TableCellBorders();
            TopBorder topBorder407 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder440 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder275 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder440 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders437.Append(topBorder407);
            tableCellBorders437.Append(leftBorder440);
            tableCellBorders437.Append(bottomBorder275);
            tableCellBorders437.Append(rightBorder440);
            Shading shading437 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment82 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties437.Append(tableCellWidth437);
            tableCellProperties437.Append(verticalMerge51);
            tableCellProperties437.Append(tableCellBorders437);
            tableCellProperties437.Append(shading437);
            tableCellProperties437.Append(tableCellVerticalAlignment82);

            Paragraph paragraph479 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "0012629E", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties479 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties479 = new ParagraphMarkRunProperties();
            RunFonts runFonts614 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold13 = new Bold();
            FontSize fontSize272 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript270 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties479.Append(runFonts614);
            paragraphMarkRunProperties479.Append(bold13);
            paragraphMarkRunProperties479.Append(fontSize272);
            paragraphMarkRunProperties479.Append(fontSizeComplexScript270);

            paragraphProperties479.Append(paragraphMarkRunProperties479);

            paragraph479.Append(paragraphProperties479);

            tableCell437.Append(tableCellProperties437);
            tableCell437.Append(paragraph479);

            TableCell tableCell438 = new TableCell();

            TableCellProperties tableCellProperties438 = new TableCellProperties();
            TableCellWidth tableCellWidth438 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders438 = new TableCellBorders();
            TopBorder topBorder408 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder441 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder276 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder441 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders438.Append(topBorder408);
            tableCellBorders438.Append(leftBorder441);
            tableCellBorders438.Append(bottomBorder276);
            tableCellBorders438.Append(rightBorder441);
            Shading shading438 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment83 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties438.Append(tableCellWidth438);
            tableCellProperties438.Append(tableCellBorders438);
            tableCellProperties438.Append(shading438);
            tableCellProperties438.Append(tableCellVerticalAlignment83);

            Paragraph paragraph480 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties480 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties480 = new ParagraphMarkRunProperties();
            RunFonts runFonts615 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Highlight highlight38 = new Highlight() { Val = HighlightColorValues.Yellow };

            paragraphMarkRunProperties480.Append(runFonts615);
            paragraphMarkRunProperties480.Append(highlight38);

            paragraphProperties480.Append(paragraphMarkRunProperties480);

            paragraph480.Append(paragraphProperties480);

            tableCell438.Append(tableCellProperties438);
            tableCell438.Append(paragraph480);

            TableCell tableCell439 = new TableCell();

            TableCellProperties tableCellProperties439 = new TableCellProperties();
            TableCellWidth tableCellWidth439 = new TableCellWidth() { Width = "2520", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders439 = new TableCellBorders();
            TopBorder topBorder409 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder442 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder277 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            RightBorder rightBorder442 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders439.Append(topBorder409);
            tableCellBorders439.Append(leftBorder442);
            tableCellBorders439.Append(bottomBorder277);
            tableCellBorders439.Append(rightBorder442);
            Shading shading439 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment84 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties439.Append(tableCellWidth439);
            tableCellProperties439.Append(tableCellBorders439);
            tableCellProperties439.Append(shading439);
            tableCellProperties439.Append(tableCellVerticalAlignment84);

            Paragraph paragraph481 = new Paragraph() { RsidParagraphMarkRevision = "00322ED6", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "009B7D94", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties481 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties481 = new ParagraphMarkRunProperties();
            RunFonts runFonts616 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold14 = new Bold();
            FontSize fontSize273 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript271 = new FontSizeComplexScript() { Val = "22" };
            Highlight highlight39 = new Highlight() { Val = HighlightColorValues.Yellow };

            paragraphMarkRunProperties481.Append(runFonts616);
            paragraphMarkRunProperties481.Append(bold14);
            paragraphMarkRunProperties481.Append(fontSize273);
            paragraphMarkRunProperties481.Append(fontSizeComplexScript271);
            paragraphMarkRunProperties481.Append(highlight39);

            paragraphProperties481.Append(paragraphMarkRunProperties481);

            Run run137 = new Run() { RsidRunProperties = "00322ED6" };

            RunProperties runProperties137 = new RunProperties();
            RunFonts runFonts617 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold15 = new Bold();
            FontSize fontSize274 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript272 = new FontSizeComplexScript() { Val = "22" };

            runProperties137.Append(runFonts617);
            runProperties137.Append(bold15);
            runProperties137.Append(fontSize274);
            runProperties137.Append(fontSizeComplexScript272);
            Text text137 = new Text();
            text137.Text = "Salinities Read By";

            run137.Append(runProperties137);
            run137.Append(text137);

            paragraph481.Append(paragraphProperties481);
            paragraph481.Append(run137);

            tableCell439.Append(tableCellProperties439);
            tableCell439.Append(paragraph481);

            TableCell tableCell440 = new TableCell();

            TableCellProperties tableCellProperties440 = new TableCellProperties();
            TableCellWidth tableCellWidth440 = new TableCellWidth() { Width = "3780", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders440 = new TableCellBorders();
            TopBorder topBorder410 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder443 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder278 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder443 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders440.Append(topBorder410);
            tableCellBorders440.Append(leftBorder443);
            tableCellBorders440.Append(bottomBorder278);
            tableCellBorders440.Append(rightBorder443);
            Shading shading440 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment85 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties440.Append(tableCellWidth440);
            tableCellProperties440.Append(tableCellBorders440);
            tableCellProperties440.Append(shading440);
            tableCellProperties440.Append(tableCellVerticalAlignment85);

            Paragraph paragraph482 = new Paragraph() { RsidParagraphMarkRevision = "00322ED6", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "009B7D94", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties482 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties482 = new ParagraphMarkRunProperties();
            RunFonts runFonts618 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color37 = new Color() { Val = "C0C0C0" };
            FontSize fontSize275 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript273 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties482.Append(runFonts618);
            paragraphMarkRunProperties482.Append(color37);
            paragraphMarkRunProperties482.Append(fontSize275);
            paragraphMarkRunProperties482.Append(fontSizeComplexScript273);

            paragraphProperties482.Append(paragraphMarkRunProperties482);

            Run run138 = new Run();

            RunProperties runProperties138 = new RunProperties();
            RunFonts runFonts619 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize276 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript274 = new FontSizeComplexScript() { Val = "20" };

            runProperties138.Append(runFonts619);
            runProperties138.Append(fontSize276);
            runProperties138.Append(fontSizeComplexScript274);
            Text text138 = new Text();
            text138.Text = "";
            if (!(string.IsNullOrWhiteSpace(labSheetA1Sheet.SalinitiesReadYear) || string.IsNullOrWhiteSpace(labSheetA1Sheet.SalinitiesReadMonth) || string.IsNullOrWhiteSpace(labSheetA1Sheet.SalinitiesReadDay)))
            {
                if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.SalinitiesReadBy))
                {
                    DateTime dateTimeSalinities = new DateTime(int.Parse(labSheetA1Sheet.SalinitiesReadYear), int.Parse(labSheetA1Sheet.SalinitiesReadMonth), int.Parse(labSheetA1Sheet.SalinitiesReadDay));
                    text138.Text = labSheetA1Sheet.SalinitiesReadBy + "        " + dateTimeSalinities.ToString("yyyy MMMM dd");
                }
            }

            run138.Append(runProperties138);
            run138.Append(text138);

            paragraph482.Append(paragraphProperties482);
            paragraph482.Append(run138);

            tableCell440.Append(tableCellProperties440);
            tableCell440.Append(paragraph482);

            tableRow28.Append(tableRowProperties25);
            tableRow28.Append(tableCell437);
            tableRow28.Append(tableCell438);
            tableRow28.Append(tableCell439);
            tableRow28.Append(tableCell440);

            TableRow tableRow29 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "009A6852", RsidTableRowProperties = "00ED7FB2" };

            TableRowProperties tableRowProperties26 = new TableRowProperties();
            TableRowHeight tableRowHeight26 = new TableRowHeight() { Val = (UInt32Value)271U };

            tableRowProperties26.Append(tableRowHeight26);

            TableCell tableCell441 = new TableCell();

            TableCellProperties tableCellProperties441 = new TableCellProperties();
            TableCellWidth tableCellWidth441 = new TableCellWidth() { Width = "8224", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge52 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders441 = new TableCellBorders();
            TopBorder topBorder411 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder444 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder279 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder444 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders441.Append(topBorder411);
            tableCellBorders441.Append(leftBorder444);
            tableCellBorders441.Append(bottomBorder279);
            tableCellBorders441.Append(rightBorder444);
            Shading shading441 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment86 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties441.Append(tableCellWidth441);
            tableCellProperties441.Append(verticalMerge52);
            tableCellProperties441.Append(tableCellBorders441);
            tableCellProperties441.Append(shading441);
            tableCellProperties441.Append(tableCellVerticalAlignment86);

            Paragraph paragraph483 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties483 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties483 = new ParagraphMarkRunProperties();
            RunFonts runFonts620 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties483.Append(runFonts620);

            paragraphProperties483.Append(paragraphMarkRunProperties483);

            Run run139 = new Run() { RsidRunProperties = "00322ED6" };

            RunProperties runProperties139 = new RunProperties();
            RunFonts runFonts621 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold16 = new Bold();
            FontSize fontSize277 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript275 = new FontSizeComplexScript() { Val = "22" };

            runProperties139.Append(runFonts621);
            runProperties139.Append(bold16);
            runProperties139.Append(fontSize277);
            runProperties139.Append(fontSizeComplexScript275);
            Text text139 = new Text();
            text139.Text = "Comments/ Non Conformances:";

            run139.Append(runProperties139);
            run139.Append(text139);

            Run run140 = new Run() { RsidRunProperties = "00322ED6", RsidRunAddition = "0012629E" };

            RunProperties runProperties140 = new RunProperties();
            RunFonts runFonts622 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold17 = new Bold();
            FontSize fontSize278 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript276 = new FontSizeComplexScript() { Val = "22" };

            runProperties140.Append(runFonts622);
            runProperties140.Append(bold17);
            runProperties140.Append(fontSize278);
            runProperties140.Append(fontSizeComplexScript276);
            Text text140 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text140.Text = "  ";

            run140.Append(runProperties140);
            run140.Append(text140);
            ProofError proofError57 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run141 = new Run() { RsidRunAddition = "00030688" };

            RunProperties runProperties141 = new RunProperties();
            RunFonts runFonts623 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            //Bold bold18 = new Bold();
            FontSize fontSize279 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript277 = new FontSizeComplexScript() { Val = "22" };

            runProperties141.Append(runFonts623);
            //runProperties141.Append(bold18);
            runProperties141.Append(fontSize279);
            runProperties141.Append(fontSizeComplexScript277);
            Text text141 = new Text();
            StringBuilder sbRunComment = new StringBuilder();
            sbRunComment.Append(labSheetA1Sheet.RunComment + " ");

            foreach (CSSPModelsDLL.Models.LabSheetA1Measurement labSheetA1Measurement in labSheetA1Sheet.LabSheetA1MeasurementList)
            {
                if (!string.IsNullOrWhiteSpace(labSheetA1Measurement.SiteComment))
                {
                    sbRunComment.Append(" [" + labSheetA1Measurement.Site + " - " + labSheetA1Measurement.SiteComment + "]");
                }
            }
            string FullRunComment = sbRunComment.ToString();

            text141.Text = (FullRunComment.Length > 250 ? FullRunComment.Substring(0, 250) : FullRunComment);

            run141.Append(runProperties141);
            run141.Append(text141);
            ProofError proofError58 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph483.Append(paragraphProperties483);
            paragraph483.Append(run139);
            paragraph483.Append(run140);
            paragraph483.Append(proofError57);
            paragraph483.Append(run141);
            paragraph483.Append(proofError58);

            tableCell441.Append(tableCellProperties441);
            tableCell441.Append(paragraph483);

            TableCell tableCell442 = new TableCell();

            TableCellProperties tableCellProperties442 = new TableCellProperties();
            TableCellWidth tableCellWidth442 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge53 = new VerticalMerge();

            TableCellBorders tableCellBorders442 = new TableCellBorders();
            TopBorder topBorder412 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder445 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder280 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder445 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders442.Append(topBorder412);
            tableCellBorders442.Append(leftBorder445);
            tableCellBorders442.Append(bottomBorder280);
            tableCellBorders442.Append(rightBorder445);
            Shading shading442 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment87 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties442.Append(tableCellWidth442);
            tableCellProperties442.Append(verticalMerge53);
            tableCellProperties442.Append(tableCellBorders442);
            tableCellProperties442.Append(shading442);
            tableCellProperties442.Append(tableCellVerticalAlignment87);

            Paragraph paragraph484 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties484 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties484 = new ParagraphMarkRunProperties();
            RunFonts runFonts624 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties484.Append(runFonts624);

            paragraphProperties484.Append(paragraphMarkRunProperties484);

            paragraph484.Append(paragraphProperties484);

            tableCell442.Append(tableCellProperties442);
            tableCell442.Append(paragraph484);

            TableCell tableCell443 = new TableCell();

            TableCellProperties tableCellProperties443 = new TableCellProperties();
            TableCellWidth tableCellWidth443 = new TableCellWidth() { Width = "2520", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders443 = new TableCellBorders();
            TopBorder topBorder413 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            LeftBorder leftBorder446 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder281 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder446 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };

            tableCellBorders443.Append(topBorder413);
            tableCellBorders443.Append(leftBorder446);
            tableCellBorders443.Append(bottomBorder281);
            tableCellBorders443.Append(rightBorder446);
            Shading shading443 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment88 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties443.Append(tableCellWidth443);
            tableCellProperties443.Append(tableCellBorders443);
            tableCellProperties443.Append(shading443);
            tableCellProperties443.Append(tableCellVerticalAlignment88);

            Paragraph paragraph485 = new Paragraph() { RsidParagraphMarkRevision = "00322ED6", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "009B7D94", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties485 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties485 = new ParagraphMarkRunProperties();
            RunFonts runFonts625 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold19 = new Bold();
            FontSize fontSize280 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript278 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties485.Append(runFonts625);
            paragraphMarkRunProperties485.Append(bold19);
            paragraphMarkRunProperties485.Append(fontSize280);
            paragraphMarkRunProperties485.Append(fontSizeComplexScript278);

            paragraphProperties485.Append(paragraphMarkRunProperties485);

            Run run142 = new Run() { RsidRunProperties = "00322ED6" };

            RunProperties runProperties142 = new RunProperties();
            RunFonts runFonts626 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold20 = new Bold();
            FontSize fontSize281 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript279 = new FontSizeComplexScript() { Val = "22" };

            runProperties142.Append(runFonts626);
            runProperties142.Append(bold20);
            runProperties142.Append(fontSize281);
            runProperties142.Append(fontSizeComplexScript279);
            Text text142 = new Text();
            text142.Text = "Results Read By";

            run142.Append(runProperties142);
            run142.Append(text142);

            paragraph485.Append(paragraphProperties485);
            paragraph485.Append(run142);

            tableCell443.Append(tableCellProperties443);
            tableCell443.Append(paragraph485);

            TableCell tableCell444 = new TableCell();

            TableCellProperties tableCellProperties444 = new TableCellProperties();
            TableCellWidth tableCellWidth444 = new TableCellWidth() { Width = "3780", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders444 = new TableCellBorders();
            TopBorder topBorder414 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder447 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)2U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder282 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder447 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders444.Append(topBorder414);
            tableCellBorders444.Append(leftBorder447);
            tableCellBorders444.Append(bottomBorder282);
            tableCellBorders444.Append(rightBorder447);
            Shading shading444 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment89 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties444.Append(tableCellWidth444);
            tableCellProperties444.Append(tableCellBorders444);
            tableCellProperties444.Append(shading444);
            tableCellProperties444.Append(tableCellVerticalAlignment89);

            Paragraph paragraph486 = new Paragraph() { RsidParagraphMarkRevision = "00322ED6", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "009B7D94", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties486 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties486 = new ParagraphMarkRunProperties();
            RunFonts runFonts627 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color38 = new Color() { Val = "C0C0C0" };
            FontSize fontSize282 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript280 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties486.Append(runFonts627);
            paragraphMarkRunProperties486.Append(color38);
            paragraphMarkRunProperties486.Append(fontSize282);
            paragraphMarkRunProperties486.Append(fontSizeComplexScript280);

            paragraphProperties486.Append(paragraphMarkRunProperties486);

            Run run143 = new Run();

            RunProperties runProperties143 = new RunProperties();
            RunFonts runFonts628 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize283 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript281 = new FontSizeComplexScript() { Val = "20" };

            runProperties143.Append(runFonts628);
            runProperties143.Append(fontSize283);
            runProperties143.Append(fontSizeComplexScript281);
            Text text143 = new Text();
            text143.Text = "";
            if (!(string.IsNullOrWhiteSpace(labSheetA1Sheet.ResultsReadYear) || string.IsNullOrWhiteSpace(labSheetA1Sheet.ResultsReadMonth) || string.IsNullOrWhiteSpace(labSheetA1Sheet.ResultsReadDay)))
            {
                if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.ResultsReadBy))
                {
                    DateTime dateTimeResultsRead = new DateTime(int.Parse(labSheetA1Sheet.ResultsReadYear), int.Parse(labSheetA1Sheet.ResultsReadMonth), int.Parse(labSheetA1Sheet.ResultsReadDay));
                    text143.Text = labSheetA1Sheet.ResultsReadBy + "        " + dateTimeResultsRead.ToString("yyyy MMMM dd");
                }
            }
            run143.Append(runProperties143);
            run143.Append(text143);

            paragraph486.Append(paragraphProperties486);
            paragraph486.Append(run143);

            tableCell444.Append(tableCellProperties444);
            tableCell444.Append(paragraph486);

            tableRow29.Append(tableRowProperties26);
            tableRow29.Append(tableCell441);
            tableRow29.Append(tableCell442);
            tableRow29.Append(tableCell443);
            tableRow29.Append(tableCell444);

            TableRow tableRow30 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "009A6852", RsidTableRowProperties = "00030688" };

            TableRowProperties tableRowProperties27 = new TableRowProperties();
            TableRowHeight tableRowHeight27 = new TableRowHeight() { Val = (UInt32Value)290U };

            tableRowProperties27.Append(tableRowHeight27);

            TableCell tableCell445 = new TableCell();

            TableCellProperties tableCellProperties445 = new TableCellProperties();
            TableCellWidth tableCellWidth445 = new TableCellWidth() { Width = "8224", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge54 = new VerticalMerge();

            TableCellBorders tableCellBorders445 = new TableCellBorders();
            TopBorder topBorder415 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder448 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder283 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder448 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders445.Append(topBorder415);
            tableCellBorders445.Append(leftBorder448);
            tableCellBorders445.Append(bottomBorder283);
            tableCellBorders445.Append(rightBorder448);
            Shading shading445 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment90 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties445.Append(tableCellWidth445);
            tableCellProperties445.Append(verticalMerge54);
            tableCellProperties445.Append(tableCellBorders445);
            tableCellProperties445.Append(shading445);
            tableCellProperties445.Append(tableCellVerticalAlignment90);

            Paragraph paragraph487 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties487 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties487 = new ParagraphMarkRunProperties();
            RunFonts runFonts629 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties487.Append(runFonts629);

            paragraphProperties487.Append(paragraphMarkRunProperties487);

            paragraph487.Append(paragraphProperties487);

            tableCell445.Append(tableCellProperties445);
            tableCell445.Append(paragraph487);

            TableCell tableCell446 = new TableCell();

            TableCellProperties tableCellProperties446 = new TableCellProperties();
            TableCellWidth tableCellWidth446 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders446 = new TableCellBorders();
            TopBorder topBorder416 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder449 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder284 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder449 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders446.Append(topBorder416);
            tableCellBorders446.Append(leftBorder449);
            tableCellBorders446.Append(bottomBorder284);
            tableCellBorders446.Append(rightBorder449);
            Shading shading446 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment91 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties446.Append(tableCellWidth446);
            tableCellProperties446.Append(tableCellBorders446);
            tableCellProperties446.Append(shading446);
            tableCellProperties446.Append(tableCellVerticalAlignment91);

            Paragraph paragraph488 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties488 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties488 = new ParagraphMarkRunProperties();
            RunFonts runFonts630 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties488.Append(runFonts630);

            paragraphProperties488.Append(paragraphMarkRunProperties488);

            paragraph488.Append(paragraphProperties488);

            tableCell446.Append(tableCellProperties446);
            tableCell446.Append(paragraph488);

            TableCell tableCell447 = new TableCell();

            TableCellProperties tableCellProperties447 = new TableCellProperties();
            TableCellWidth tableCellWidth447 = new TableCellWidth() { Width = "2520", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders447 = new TableCellBorders();
            TopBorder topBorder417 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder450 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder285 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder450 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders447.Append(topBorder417);
            tableCellBorders447.Append(leftBorder450);
            tableCellBorders447.Append(bottomBorder285);
            tableCellBorders447.Append(rightBorder450);
            Shading shading447 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment92 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties447.Append(tableCellWidth447);
            tableCellProperties447.Append(tableCellBorders447);
            tableCellProperties447.Append(shading447);
            tableCellProperties447.Append(tableCellVerticalAlignment92);

            Paragraph paragraph489 = new Paragraph() { RsidParagraphMarkRevision = "00322ED6", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "009B7D94", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties489 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties489 = new ParagraphMarkRunProperties();
            RunFonts runFonts631 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold21 = new Bold();
            FontSize fontSize284 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript282 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties489.Append(runFonts631);
            paragraphMarkRunProperties489.Append(bold21);
            paragraphMarkRunProperties489.Append(fontSize284);
            paragraphMarkRunProperties489.Append(fontSizeComplexScript282);

            paragraphProperties489.Append(paragraphMarkRunProperties489);

            Run run144 = new Run() { RsidRunProperties = "006B4849" };

            RunProperties runProperties144 = new RunProperties();
            RunFonts runFonts632 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold22 = new Bold();
            FontSize fontSize285 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript283 = new FontSizeComplexScript() { Val = "22" };

            runProperties144.Append(runFonts632);
            runProperties144.Append(bold22);
            runProperties144.Append(fontSize285);
            runProperties144.Append(fontSizeComplexScript283);
            Text text144 = new Text();
            text144.Text = "Results Recorded By";

            run144.Append(runProperties144);
            run144.Append(text144);

            paragraph489.Append(paragraphProperties489);
            paragraph489.Append(run144);

            tableCell447.Append(tableCellProperties447);
            tableCell447.Append(paragraph489);

            TableCell tableCell448 = new TableCell();

            TableCellProperties tableCellProperties448 = new TableCellProperties();
            TableCellWidth tableCellWidth448 = new TableCellWidth() { Width = "3780", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders448 = new TableCellBorders();
            TopBorder topBorder418 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder451 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder286 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder451 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders448.Append(topBorder418);
            tableCellBorders448.Append(leftBorder451);
            tableCellBorders448.Append(bottomBorder286);
            tableCellBorders448.Append(rightBorder451);
            Shading shading448 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment93 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties448.Append(tableCellWidth448);
            tableCellProperties448.Append(tableCellBorders448);
            tableCellProperties448.Append(shading448);
            tableCellProperties448.Append(tableCellVerticalAlignment93);

            Paragraph paragraph490 = new Paragraph() { RsidParagraphMarkRevision = "00322ED6", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "0012629E", RsidRunAdditionDefault = "00030688" };

            ParagraphProperties paragraphProperties490 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties490 = new ParagraphMarkRunProperties();
            RunFonts runFonts633 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color39 = new Color() { Val = "C0C0C0" };
            FontSize fontSize286 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript284 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties490.Append(runFonts633);
            paragraphMarkRunProperties490.Append(color39);
            paragraphMarkRunProperties490.Append(fontSize286);
            paragraphMarkRunProperties490.Append(fontSizeComplexScript284);

            paragraphProperties490.Append(paragraphMarkRunProperties490);

            Run run145 = new Run();

            RunProperties runProperties145 = new RunProperties();
            RunFonts runFonts634 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize287 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript285 = new FontSizeComplexScript() { Val = "20" };

            runProperties145.Append(runFonts634);
            runProperties145.Append(fontSize287);
            runProperties145.Append(fontSizeComplexScript285);
            Text text145 = new Text();
            text145.Text = "";
            if (!(string.IsNullOrWhiteSpace(labSheetA1Sheet.ResultsRecordedYear) || string.IsNullOrWhiteSpace(labSheetA1Sheet.ResultsRecordedMonth) || string.IsNullOrWhiteSpace(labSheetA1Sheet.ResultsRecordedDay)))
            {
                if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.ResultsRecordedBy))
                {
                    DateTime dateTimeResultsRecorded = new DateTime(int.Parse(labSheetA1Sheet.ResultsRecordedYear), int.Parse(labSheetA1Sheet.ResultsRecordedMonth), int.Parse(labSheetA1Sheet.ResultsRecordedDay));
                    text145.Text = labSheetA1Sheet.ResultsRecordedBy + "        " + dateTimeResultsRecorded.ToString("yyyy MMMM dd");
                }
            }

            run145.Append(runProperties145);
            run145.Append(text145);

            paragraph490.Append(paragraphProperties490);
            paragraph490.Append(run145);

            tableCell448.Append(tableCellProperties448);
            tableCell448.Append(paragraph490);

            tableRow30.Append(tableRowProperties27);
            tableRow30.Append(tableCell445);
            tableRow30.Append(tableCell446);
            tableRow30.Append(tableCell447);
            tableRow30.Append(tableCell448);

            TableRow tableRow31 = new TableRow() { RsidTableRowMarkRevision = "00ED7FB2", RsidTableRowAddition = "009A6852", RsidTableRowProperties = "00322ED6" };

            TableRowProperties tableRowProperties28 = new TableRowProperties();
            TableRowHeight tableRowHeight28 = new TableRowHeight() { Val = (UInt32Value)408U };

            tableRowProperties28.Append(tableRowHeight28);

            TableCell tableCell449 = new TableCell();

            TableCellProperties tableCellProperties449 = new TableCellProperties();
            TableCellWidth tableCellWidth449 = new TableCellWidth() { Width = "8224", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge55 = new VerticalMerge();

            TableCellBorders tableCellBorders449 = new TableCellBorders();
            TopBorder topBorder419 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder452 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder287 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder452 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders449.Append(topBorder419);
            tableCellBorders449.Append(leftBorder452);
            tableCellBorders449.Append(bottomBorder287);
            tableCellBorders449.Append(rightBorder452);
            Shading shading449 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment94 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties449.Append(tableCellWidth449);
            tableCellProperties449.Append(verticalMerge55);
            tableCellProperties449.Append(tableCellBorders449);
            tableCellProperties449.Append(shading449);
            tableCellProperties449.Append(tableCellVerticalAlignment94);

            Paragraph paragraph491 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties491 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties491 = new ParagraphMarkRunProperties();
            RunFonts runFonts635 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize288 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript286 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties491.Append(runFonts635);
            paragraphMarkRunProperties491.Append(fontSize288);
            paragraphMarkRunProperties491.Append(fontSizeComplexScript286);

            paragraphProperties491.Append(paragraphMarkRunProperties491);

            paragraph491.Append(paragraphProperties491);

            tableCell449.Append(tableCellProperties449);
            tableCell449.Append(paragraph491);

            TableCell tableCell450 = new TableCell();

            TableCellProperties tableCellProperties450 = new TableCellProperties();
            TableCellWidth tableCellWidth450 = new TableCellWidth() { Width = "236", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders450 = new TableCellBorders();
            TopBorder topBorder420 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder453 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder288 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder453 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders450.Append(topBorder420);
            tableCellBorders450.Append(leftBorder453);
            tableCellBorders450.Append(bottomBorder288);
            tableCellBorders450.Append(rightBorder453);
            Shading shading450 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment95 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties450.Append(tableCellWidth450);
            tableCellProperties450.Append(tableCellBorders450);
            tableCellProperties450.Append(shading450);
            tableCellProperties450.Append(tableCellVerticalAlignment95);

            Paragraph paragraph492 = new Paragraph() { RsidParagraphMarkRevision = "00ED7FB2", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00263D93", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties492 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties492 = new ParagraphMarkRunProperties();
            RunFonts runFonts636 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };

            paragraphMarkRunProperties492.Append(runFonts636);

            paragraphProperties492.Append(paragraphMarkRunProperties492);

            paragraph492.Append(paragraphProperties492);

            tableCell450.Append(tableCellProperties450);
            tableCell450.Append(paragraph492);

            TableCell tableCell451 = new TableCell();

            TableCellProperties tableCellProperties451 = new TableCellProperties();
            TableCellWidth tableCellWidth451 = new TableCellWidth() { Width = "2520", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders451 = new TableCellBorders();
            TopBorder topBorder421 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder454 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder289 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder454 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders451.Append(topBorder421);
            tableCellBorders451.Append(leftBorder454);
            tableCellBorders451.Append(bottomBorder289);
            tableCellBorders451.Append(rightBorder454);
            Shading shading451 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment96 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties451.Append(tableCellWidth451);
            tableCellProperties451.Append(tableCellBorders451);
            tableCellProperties451.Append(shading451);
            tableCellProperties451.Append(tableCellVerticalAlignment96);

            Paragraph paragraph493 = new Paragraph() { RsidParagraphMarkRevision = "00322ED6", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "006B4849", RsidRunAdditionDefault = "009A6852" };

            ParagraphProperties paragraphProperties493 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties493 = new ParagraphMarkRunProperties();
            RunFonts runFonts637 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold23 = new Bold();
            FontSize fontSize289 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript287 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties493.Append(runFonts637);
            paragraphMarkRunProperties493.Append(bold23);
            paragraphMarkRunProperties493.Append(fontSize289);
            paragraphMarkRunProperties493.Append(fontSizeComplexScript287);

            paragraphProperties493.Append(paragraphMarkRunProperties493);

            Run run146 = new Run();

            RunProperties runProperties146 = new RunProperties();
            RunFonts runFonts638 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold24 = new Bold();
            FontSize fontSize290 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript288 = new FontSizeComplexScript() { Val = "22" };
            Highlight highlight2 = new Highlight() { Val = HighlightColorValues.Yellow };

            runProperties146.Append(runFonts638);
            runProperties146.Append(bold24);
            runProperties146.Append(fontSize290);
            runProperties146.Append(fontSizeComplexScript288);
            runProperties146.Append(highlight2);
            Text text146 = new Text();
            text146.Text = "Supervisor Approval";

            run146.Append(runProperties146);
            run146.Append(text146);

            paragraph493.Append(paragraphProperties493);
            paragraph493.Append(run146);

            tableCell451.Append(tableCellProperties451);
            tableCell451.Append(paragraph493);
            TableCell tableCell452 = new TableCell();

            TableCellProperties tableCellProperties452 = new TableCellProperties();
            TableCellWidth tableCellWidth452 = new TableCellWidth() { Width = "3780", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders452 = new TableCellBorders();
            TopBorder topBorder422 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder455 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder290 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder455 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders452.Append(topBorder422);
            tableCellBorders452.Append(leftBorder455);
            tableCellBorders452.Append(bottomBorder290);
            tableCellBorders452.Append(rightBorder455);
            Shading shading452 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment97 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties452.Append(tableCellWidth452);
            tableCellProperties452.Append(tableCellBorders452);
            tableCellProperties452.Append(shading452);
            tableCellProperties452.Append(tableCellVerticalAlignment97);

            Paragraph paragraph494 = new Paragraph() { RsidParagraphMarkRevision = "00322ED6", RsidParagraphAddition = "009A6852", RsidParagraphProperties = "00AB7586", RsidRunAdditionDefault = "00AB7586" };

            ParagraphProperties paragraphProperties494 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties494 = new ParagraphMarkRunProperties();
            RunFonts runFonts640 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize292 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript290 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties494.Append(runFonts640);
            paragraphMarkRunProperties494.Append(fontSize292);
            paragraphMarkRunProperties494.Append(fontSizeComplexScript290);

            paragraphProperties494.Append(paragraphMarkRunProperties494);

            Run run148 = new Run();

            RunProperties runProperties148 = new RunProperties();
            RunFonts runFonts641 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize293 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript291 = new FontSizeComplexScript() { Val = "20" };

            runProperties148.Append(runFonts641);
            runProperties148.Append(fontSize293);
            runProperties148.Append(fontSizeComplexScript291);
            Text text148 = new Text();
            text148.Text = "";
            if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.ApprovedBySupervisorInitials))
            {
                if (!(string.IsNullOrWhiteSpace(labSheetA1Sheet.ApprovalYear) || string.IsNullOrWhiteSpace(labSheetA1Sheet.ApprovalMonth) || string.IsNullOrWhiteSpace(labSheetA1Sheet.ApprovalDay)))
                {
                    if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.ApprovedBySupervisorInitials))
                    {
                        DateTime dateTimeApproavalDate = new DateTime(int.Parse(labSheetA1Sheet.ApprovalYear), int.Parse(labSheetA1Sheet.ApprovalMonth), int.Parse(labSheetA1Sheet.ApprovalDay));
                        text148.Text = labSheetA1Sheet.ApprovedBySupervisorInitials + "        " + dateTimeApproavalDate.ToString("yyyy MMMM dd");
                    }
                }
            }

            run148.Append(runProperties148);
            run148.Append(text148);

            paragraph494.Append(paragraphProperties494);
            paragraph494.Append(run148);

            tableCell452.Append(tableCellProperties452);
            tableCell452.Append(paragraph494);

            tableRow31.Append(tableRowProperties28);
            tableRow31.Append(tableCell449);
            tableRow31.Append(tableCell450);
            tableRow31.Append(tableCell451);
            tableRow31.Append(tableCell452);

            table4.Append(tableProperties4);
            table4.Append(tableGrid4);
            table4.Append(tableRow27);
            table4.Append(tableRow28);
            table4.Append(tableRow29);
            table4.Append(tableRow30);
            table4.Append(tableRow31);
        }
    }
}
