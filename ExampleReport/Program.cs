using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace ExampleReport
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelReport report = new ExcelReport();
            report.Initialise();

            report.Append(1, 2, 3, 4, 5, 6);
            report.Append(1, 2, 3, 4, 5, 6);
            report.Append(1, 2, 3, 4, 5, 6);
            report.Append(1, 2, 3, 4, 5, 6);

            List<String> errors = new List<String>();
            errors.Add("error1");
            errors.Add("error2");
            errors.Add("error3");
            report.Append(1, 2, null, errors);
            report.Append(1, 2, "expression", errors);
            report.Append(1, 2, "expression", errors);
            errors.Add("error4");
            errors.Add("error5");
            report.Append(3, 4, "expression", errors);
            report.Append(3, 4, "expression", errors);

            String filename = report.Export("Example.xml", true);

			if (MessageBox.Show("Would you want open the report?", "ExcelReport Example", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
			{
				Process proccess = new Process();
				proccess.StartInfo = new ProcessStartInfo(filename);
				proccess.Start();
			}
        }
    }
}
