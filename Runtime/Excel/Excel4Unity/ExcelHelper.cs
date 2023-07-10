using System;
using UnityEngine;
using System.Collections;
using OfficeOpenXml;
using System.IO;
using ExcelDataReader;
using System.Data;


public class ExcelHelper
{

    public static Excel LoadExcel(string path)
    {
        try
        {
            FileInfo file = new FileInfo(path);
            ExcelPackage ep = new ExcelPackage(file);
            Excel xls = new Excel(ep.Workbook, ep);
            return xls;
        }
        catch (System.Exception ex)
        {
            UnityEditor.EditorUtility.DisplayDialog("Error", "Excel file is open, close and retry!\r\n" + path, "OK", "");
        }

        return null;
    }

	public static Excel CreateExcel(string path) {
		ExcelPackage ep = new ExcelPackage ();
		ep.Workbook.Worksheets.Add ("sheet");
		Excel xls = new Excel(ep.Workbook, ep);
		SaveExcel (xls, path);
		return xls;
	}

    public static void SaveExcel(Excel xls, string path)
    {
        FileInfo output = new FileInfo(path);
        using (ExcelPackage ep = new ExcelPackage())
        {
            for (int i = 0; i < xls.Tables.Count; i++)
            {
                ExcelTable table = xls.Tables[i];
                ExcelWorksheet sheet = ep.Workbook.Worksheets.Add(table.TableName);
                for (int row = 1; row <= table.NumberOfRows; row++)
                {
                    for (int column = 1; column <= table.NumberOfColumns; column++)
                    {
                        sheet.Cells[row, column].Value = table.GetValue(row, column);
                    }
                }
            }
            ep.SaveAs(output);
        }
    }

    public static void SaveCSV(Excel xls, string path)
    {        
        for (int i = 0; i < xls.Tables.Count; i++)
        {
            string str = "";
            ExcelTable table = xls.Tables[i];
            for (int row = 1; row <= table.NumberOfRows; row++)
            {
                for (int column = 1; column <= table.NumberOfColumns; column++)
                {
                    str += "\"" + table.GetValue(row, column).ToString() + "\"";
                    if (column <= table.NumberOfColumns)
                    {
                        str += ",";
                    }
                }

                if (row <= table.NumberOfRows)
                {
                    str += "\r\n";
                }
            }

            using (var stream = System.IO.File.CreateText(path))
            {
                stream.Write(str);
                stream.Close();
            }

            break;
        }
    }

    public static DataTable DecodeExcel(string path)
    {
        if(path.Contains("~$"))
        {
            return null;
        }

        FileStream fStream = File.Open(path, FileMode.Open, FileAccess.Read);
        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fStream);

        DataSet result = excelReader.AsDataSet();
        if (result.Tables.Count <= 0)
        {
            Debug.LogError("空表");
            excelReader.Close();
            return null;
        }
        DataTable table = result.Tables[0];
        excelReader.Close();

        return table;
    }
}
