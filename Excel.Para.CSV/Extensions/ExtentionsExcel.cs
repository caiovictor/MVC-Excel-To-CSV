using Excel;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace Excel.Para.CSV.Extensions
{
    public static class ExtentionsExcel
    {
        public static DataSet LerArquivoExcel(string file)
        {
            IExcelDataReader excelReader;
            DataSet dsExcel = new DataSet();

            FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
            if (file.EndsWith(".xlsx"))
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream); //.xlsx
            else
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream); //.xls

            dsExcel = excelReader.AsDataSet();
            excelReader.Close();

            return dsExcel;
        }

        public static void ConverterCSV(DataSet result, bool todasPlanilhas, string output)
        {
            string stringCSV = "";
            int sheet_no = 0, row_no = 0;

            while (sheet_no < result.Tables.Count)
            {
                //var planilha = result.Tables[0].TableName.ToString();
                while (row_no < result.Tables[sheet_no].Rows.Count)
                {
                    for (int i = 0; i < result.Tables[sheet_no].Columns.Count; i++)
                    {
                        stringCSV += result.Tables[sheet_no].Rows[row_no][i].ToString() + ";";
                    }
                    row_no++;
                    stringCSV += "\n";
                }

                if (!todasPlanilhas) break;

                row_no = 0;
                sheet_no++;
            }

            StreamWriter csv = new StreamWriter(@output, false);
            csv.Write(stringCSV);
            csv.Close();
        }
    }
}