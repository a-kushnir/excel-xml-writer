using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using ExcelXmlWriter;

namespace ExampleReport
{
    public class ExcelReport
    {
        #region Private Fields
        private Workbook mWorkbook;
        private Worksheet mWsDetail;
        private Worksheet mWsStatistic;
        private WorksheetStyle mStyleHeader;
        private WorksheetStyle mStyleRegularRow;
        private WorksheetStyle mStyleHighlightRow;

        private bool mHighlightDetail;
        private bool mHighlightStatistic;
        #endregion

        #region Public Methods
        public void Initialise()
        {
            mWorkbook = new Workbook();
            mWorkbook.Properties.Create("Andrew Kushnir", "10.0");

            mWsDetail = mWorkbook.NewWorksheet("Detail");
            mWsStatistic = mWorkbook.NewWorksheet("Statistic");

            InitialiseStyles();
            InitialiseTables();
            CreateHeaders();
        	SetupPrintingOptions();
        }

        public void Append(Int32 clientId, Int32 recordId, String expression, List<String> errors)
        {
            WorksheetStyle style = mHighlightDetail ? mStyleHighlightRow : mStyleRegularRow;
            mHighlightDetail = !mHighlightDetail;

            Int32 rows = errors.Count > 4 ? errors.Count : 4;

            WorksheetCell cell;
            WorksheetRow row = mWsDetail.Table.NewRow();
            cell = row.NewCell("1234567890123 SELECT *\nFROM Client\nWHERE ClientID = 10\nAND name = 'Kinetics'", DataType.String, style);

        	cell.SetBold(3, 6);
			cell.SetItalic(4, 6);
        	cell.SetColor(5, 6, Color.Yellow);
        	cell.SetFont(6, 6, "Times New Roman", 24);

            cell.MergeDown = rows - 1;
            if (String.IsNullOrEmpty(expression))
            {
                cell.MergeAcross = 1;
                //cell.MergeDown++;
            }
            else
            {
                cell = row.NewCell(expression, DataType.String, style);
                cell.MergeDown = rows - 1;
            }
            
            //mWsDetail.Table[2, 1].MergeAcross = 10;
            //mWsDetail.Table[1, 2].MergeDown = 10;

            mWsDetail.Table.Cells[1][1].Data = UInt32.MaxValue;

            bool first = true;
            foreach(String error in errors)
            {
                if (first)
                    row = mWsDetail.Table.NewRow();
                first = false;

                row.NewCell(error, DataType.String, style).Hyperlink.SetPlace(mWsStatistic.Table.Cells[0][0]);
            }

            for(Int32 i = errors.Count; i < 4; i++)
            {
                row = mWsDetail.Table.NewRow();
                row.NewCell(String.Empty, DataType.String, style);
            }
        }

        public void Append(Int32 clientId, Int32 totalFormulas, Int32 validationErrors, Int32 nolockErrors, Int32 notExistsErrors, Int32 cursorErrors)
        {
            WorksheetStyle style = mHighlightStatistic ? mStyleHighlightRow : mStyleRegularRow;
            mHighlightStatistic = !mHighlightStatistic;

            WorksheetRow row = mWsStatistic.Table.NewRow();
            row.NewCell(clientId.ToString(), DataType.Number, style);
            row.NewCell(totalFormulas.ToString(), DataType.Number, style);
            row.NewCell(validationErrors.ToString(), DataType.Number, style);
            row.NewCell(nolockErrors.ToString(), DataType.Number, style);
            row.NewCell(notExistsErrors.ToString(), DataType.Number, style);
            row.NewCell(cursorErrors.ToString(), DataType.Number, style).Comment.Create("Andrew Kushnir", "test comment\ntest comment");
        }

        public void Export(String filename)
        {
            Export(filename, false);
        }

        public String Export(String filename, bool appendDate)
        {
            if (String.IsNullOrEmpty(filename))
                throw new ArgumentException("filename");

            if (appendDate)
            {
                filename = Path.Combine(Path.GetDirectoryName(filename),
                    Path.GetFileNameWithoutExtension(filename) + DateTime.Now.ToString("yyMMddhhmmss") + Path.GetExtension(filename));
            }

            mWorkbook.Save(filename);

        	return filename;
        }
        #endregion

        #region Private Methods
        private void InitialiseStyles()
        {
            mStyleHeader = mWorkbook.NewStyle("Header", new WorksheetStyleFont("Arial", 10, true));
            mStyleHeader.Interior = new WorksheetStyleInterior(Color.Orange);
            mStyleHeader.Borders.Create(2);

            mStyleRegularRow = mWorkbook.NewStyle("RegularRow", new WorksheetStyleFont("Arial", 10));
            mStyleRegularRow.Aligment.Horizontal = StyleHorizontalAlignment.Justify;
            mStyleRegularRow.Aligment.Vertical = StyleVerticalAlignment.Top;
            mStyleRegularRow.Borders.Create(1, 2, 1, 2);

            mStyleHighlightRow = mWorkbook.NewStyle("HighlightRow", new WorksheetStyleFont("Arial", 10));
            mStyleHighlightRow.Aligment.Horizontal = StyleHorizontalAlignment.Justify;
            mStyleHighlightRow.Aligment.Vertical = StyleVerticalAlignment.Top;
            mStyleHighlightRow.Interior.Color = Color.FromArgb(204, 255, 255);
            mStyleHighlightRow.Borders.Create(1, 2, 1, 2);
        }

        private void InitialiseTables()
        {
            mWsDetail.Table.NewColumn(200);
            mWsDetail.Table.NewColumn(200);
            mWsDetail.Table.NewColumn(200);

            mWsStatistic.Table.NewColumn(50);
            mWsStatistic.Table.NewColumn(85);
            mWsStatistic.Table.NewColumn(85);
            mWsStatistic.Table.NewColumn(85);
            mWsStatistic.Table.NewColumn(85);
            mWsStatistic.Table.NewColumn(85);
        }

        private void CreateHeaders()
        {
            WorksheetRow row = mWsDetail.Table.NewRow();
            row.NewCell("Select Statement", mStyleHeader);
            row.NewCell("Formula", mStyleHeader);
            row.NewCell("Errors", mStyleHeader);

            row = mWsStatistic.Table.NewRow();
            row.NewCell("ClientID", mStyleHeader);
            row.NewCell("Total Formulas", mStyleHeader);
            row.NewCell("Validation Errors", mStyleHeader);
            row.NewCell("'Nolock' Errors", mStyleHeader);
            row.NewCell("'Not Exists' Errors", mStyleHeader);
            row.NewCell("'Cursor' Errors", mStyleHeader);
        }

	    private void SetupPrintingOptions()
	    {
	    	WorksheetOptions options = mWorkbook.Worksheets[1].Options;
	    	options.ActivePane = 3;
	    	options.FitToPage = true;
	    	options.FreezePanes = true;
	    	options.GridLineColor = "#33FFCC";
	    	options.LeftColumnRightPane = 2;
			options.PageSetup.Footer.Data = "Footer.Data";
	    	options.PageSetup.Footer.Margin = 2.3;
			options.PageSetup.Header.Data = "Header.Data";
			options.PageSetup.Header.Margin = 1.3;
	    	options.PageSetup.Layout.CenterHorizontal = true;
	    	options.PageSetup.Layout.CenterVertical = true;
	    	options.PageSetup.Layout.Orientation = Orientation.Landscape;
	    	options.PageSetup.Layout.StartPageNumber = 4;
	    	options.PageSetup.PageMargins.Bottom = 1.5;
			options.PageSetup.PageMargins.Left = 2.5;
			options.PageSetup.PageMargins.Right = 3.5;
			options.PageSetup.PageMargins.Top = 4.5;
	    	options.Print.BlackAndWhite = true;
	    	options.Print.CommentsLayout = PrintCommentsLayout.SheetEnd;
	    	options.Print.DraftQuality = true;
	    	options.Print.FitHeight = 5;
	    	options.Print.FitWidth = 6;
	    	options.Print.GridLines = true;
	    	options.Print.HorizontalResolution = 1200;
			options.Print.LeftToRight = true;
	    	options.Print.PaperSizeIndex = 1;
	    	options.Print.PrintErrors = PrintErrorsOption.Displayed;
	    	options.Print.RowColHeadings = true;
	    	options.Print.Scale = 103;
	    	options.Print.ValidPrinterInfo = true;
	    	options.Print.VerticalResolution = 1200;
	    	options.ProtectObjects = true;
	    	options.ProtectScenarios = true;
	    	options.Selected = true;
	    	options.SplitHorizontal = 3;
	    	options.SplitVertical = 3;
	    	options.TopRowBottomPane = 4;
	    	options.TopRowVisible = 5;
			// options.ViewableRange
	    }

    	#endregion
    }
}
