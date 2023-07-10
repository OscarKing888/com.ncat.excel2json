using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

#if false
using System.Runtime.InteropServices; //提供 COM互操作的库
using Excel = Microsoft.Office.Interop.Excel; 
using Office = Microsoft.Office.Core;

namespace ExcelUtil
{
	public static class StringExtensions
	{
		public static bool Contains(this string source, string toCheck, StringComparison comp)
		{
			return source?.IndexOf(toCheck, comp) >= 0;
		}
	}


	public static class ExcelUtil
    {
		static Excel.Application ExcelApp;

		public static void Init()
        {
			ExcelApp = new Excel.Application();
			ExcelApp.Visible = false;
			ExcelApp.ScreenUpdating = false;

		}

		public static void UnInit()
        {			
			ExcelApp.Quit();
			ExcelApp = null;
        }

		public static void ConvertToCSV(string srcFilePath, string destFilePath)
        {
			Microsoft.Office.Interop.Excel.Workbook workbook = ExcelApp.Workbooks.Open(srcFilePath);
			workbook.SaveAs(destFilePath, Excel.XlFileFormat.xlCSVWindows);
			workbook.Close(false);
        }
	}
}

#endif