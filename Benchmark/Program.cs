using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using Hess.ExcelXmlWriter;

namespace Benchmark
{
	class Program
    {
        #region Fields

	    private static Int64 memory;
        private static DateTime timeStart;
		private static Boolean writeToFile;
		private static Int32 fileIndex;
		private static String baseDirectory;
		private static String testDirectory;

        #endregion

        #region Program

        static void Main(string[] args)
		{
			// Allow to write generated workbooks to file.
			writeToFile = true;
			// Writes to specified directory otherwise program base directory.
			baseDirectory = @"c:\";

			Introduce();
            
			RunSimpleFillingTest(4096, 1);
            GC.Collect();
			RunSimpleFillingTest(16384, 1);
            GC.Collect();
			RunSimpleFillingTest(65536, 1);
            GC.Collect();
			RunSimpleFillingTest(4096, 8);
            GC.Collect();
			RunSimpleFillingTest(16384, 8);
            GC.Collect();
			RunSimpleFillingTest(65536, 8);
            GC.Collect();
            
            RunOrderedCellMergingTest(2048, 2, 2);
            GC.Collect();
            RunOrderedCellMergingTest(8192, 2, 2);
            GC.Collect();
            RunOrderedCellMergingTest(2048, 4, 4);
            GC.Collect();
            RunOrderedCellMergingTest(8192, 4, 4);
            GC.Collect();
            RunOrderedCellMergingTest(2048, 6, 6);
            GC.Collect();
            RunOrderedCellMergingTest(8192, 6, 6);
            GC.Collect();
            
            RunChaosCellMergingTest(4096, 1);
            GC.Collect();
            RunChaosCellMergingTest(16384, 1);
            GC.Collect();
            RunChaosCellMergingTest(512, 8);
            GC.Collect();
            RunChaosCellMergingTest(2048, 8);
            GC.Collect();

			Wait();
        }

        #endregion

        #region Tests

        private static void RunSimpleFillingTest(Int32 rows, Int32 columns)
		{
			Workbook wb = new Workbook();
			WorksheetTable table = wb.NewWorksheet().Table;

			StartTest(String.Format("Cell filling with data ({0}õ{1})", rows, columns));

			for (int i = 0; i < rows; i++)
			{
				WorksheetRow row = table.NewRow();
				for (int j = 0; j < columns; j++)
					row.NewCell("test");
			}

			EndTest();

			SaveWorksheetTest(wb);
		}

        private static void RunOrderedCellMergingTest(Int32 rows, Int32 rowPerCell, Int32 columns)
		{
			Workbook wb = new Workbook();
			WorksheetTable table = wb.NewWorksheet().Table;
            WorksheetStyle lightStyle = wb.NewStyle("l");
            lightStyle.Interior.Color = Color.LightGreen;
            lightStyle.Borders.Create(1);
            WorksheetStyle normalStyle = wb.NewStyle("n");
            normalStyle.Interior.Color = Color.LightSkyBlue;
            normalStyle.Borders.Create(1);

            StartTest(String.Format("Organizated ñell merging ({0}/R{1}C{2})", rows, rowPerCell, columns));

		    Int32 rollbacks = 0;

            Boolean light = false;
			for (int i = 0; i < rows; i++)
			{
				WorksheetRow row = table.NewRow();
                for (int j = 0; j < columns; j++)
                    row.NewCell(String.Empty, light ? lightStyle : normalStyle).MergeDown = rowPerCell - 1;

                Boolean doubleCell = false;
                for (int k = 0; k < rowPerCell; k++)
                {
                    if (k != 0)
                        row = table.NewRow();

                    for (int j = 0; j < columns; j++)
                    {

                        if (doubleCell)
                            row.NewCell(String.Empty, light ? lightStyle : normalStyle).MergeAcross = 1;
                        else
                            row.NewCell(String.Empty, light ? lightStyle : normalStyle);
    
                    }

                    doubleCell = !doubleCell;
                }

                light = !light;
			}

			EndTestRollbacks(rollbacks);

			SaveWorksheetTest(wb);
		}

        private static void RunChaosCellMergingTest(Int32 rows, Int32 columns)
		{
			Random r = new Random();
			Workbook wb = new Workbook();
			WorksheetTable table = wb.NewWorksheet().Table;
		    WorksheetStyle vertStyle = wb.NewStyle("v");
		    vertStyle.Interior.Color = Color.Green;
            vertStyle.Borders.Create(1);
            WorksheetStyle horzStyle = wb.NewStyle("h");
            horzStyle.Interior.Color = Color.DarkViolet;
            horzStyle.Borders.Create(1);

			StartTest(String.Format("Chaotic cell merging ({0}õ{1})", rows, columns));

		    Int32 rollbacks = 0;

			for (int i = 0; i < rows; i++)
			{
				WorksheetRow row = table.NewRow();
				for (int j = 0; j < columns; j++)
					if (r.Next(2) == 0)
						row.NewCell(String.Empty, vertStyle).MergeDown = r.Next(5);
					else
                        try
                        {
                            row.NewCell(String.Empty, horzStyle).MergeAcross = r.Next(5);
                        }
                        catch (ExcelXmlWriterException)
                        {
                            rollbacks++;
                        }
			}

			EndTestRollbacks(rollbacks);

			SaveWorksheetTest(wb);
		}

		private static void RunCellFormattingTest(Int32 rows, Int32 columns)
		{
		}

		private static void RunComplexTest(Int32 rows, Int32 columns)
		{
        }

        #endregion

        #region Test Helpers

        private static void SaveWorksheetTest(Workbook wb)
		{
			Console.ForegroundColor = ConsoleColor.DarkGray;

			if(writeToFile)
			{
				if (String.IsNullOrEmpty(baseDirectory))
					baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
				if (String.IsNullOrEmpty(testDirectory))
					testDirectory = Path.Combine(baseDirectory, String.Format("Hess.ExcelXmlBenchmark{0}", DateTime.Now.ToString("yyyyMMddhhmmss"))); 

				StartTest("Generating xml data to file");
			    String filename = Path.Combine(testDirectory, String.Format(@"test#{0}.xml", fileIndex));
                wb.Save(filename);
				fileIndex++;
                EndTestFileSize(new FileInfo(filename).Length);
			}
			else
			{
				StartTest("Generating xml data to memory stream");
				MemoryStream ms = new MemoryStream();
				wb.Save(ms);
				EndTestFileSize(ms.Length);
			}

			Console.ForegroundColor = ConsoleColor.White;
		}

		private static void StartTest(String description)
		{
            memory = Process.GetCurrentProcess().WorkingSet64;

			Console.Write(Expand(String.Format(" + {0}... ", description), 40));
			timeStart = DateTime.Now;
		}

        private static void EndTest()
        {
            EndTest(0, 0);
        }

        private static void EndTestFileSize(Int64 filesize)
        {
            EndTest(filesize, 0);
        }

	    private static void EndTestRollbacks(Int32 rollbacks)
        {
            EndTest(0, rollbacks);
        }

        private static void EndTest(Int64 filesize, Int32 rollbacks)
		{
			TimeSpan ts = DateTime.Now - timeStart;

			if (ts.TotalSeconds < 1)
				Console.ForegroundColor = ConsoleColor.Green;
			else if (ts.TotalSeconds < 5)
				Console.ForegroundColor = ConsoleColor.Yellow;
			else 
				Console.ForegroundColor = ConsoleColor.Red;

            Console.Write(Expand(String.Format("{0:0.00}s", ts.TotalSeconds), 8));

            Console.ForegroundColor = ConsoleColor.Blue;
            Console.Write(Expand(String.Format("{0:0.0}Mb", (Process.GetCurrentProcess().WorkingSet64 - memory) / 1048576.0), 10));

            Console.ForegroundColor = ConsoleColor.DarkGray;
            
            if (filesize > 0)
                Console.Write("{0:0.00}Mb", filesize/1048576.0);

            if (rollbacks > 0)
                Console.Write("{0} rollbacks", rollbacks);

            Console.WriteLine();
		}

		private static void Introduce()
		{
			Console.ForegroundColor = ConsoleColor.Gray;
			Console.WriteLine("=======================================");
			Console.WriteLine(" Hess.ExcelXmlWriter benchmark utility");
			Console.WriteLine("=======================================");
			Console.ForegroundColor = ConsoleColor.White;
		}

		private static void Wait()
		{
			Console.ForegroundColor = ConsoleColor.Gray;
			Console.WriteLine("=======================================");
			Console.Write("Press any key for continue");
			Console.ReadLine();
        }

        private static String Expand(String source, Int32 lenght)
        {
            if (source.Length < lenght)
                return source + new String(' ', lenght - source.Length);
            return source;
        }

	    #endregion
    }
}