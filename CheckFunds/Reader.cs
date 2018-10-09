using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace CheckFunds
{
    class Reader
    {
        private const string _marketingDirectory = @"C:\Temp\Marketing\";

        public List<string> SupList { get; } = new List<string>();

        public Reader()
        {
            FileInfo supFile = new FileInfo(_marketingDirectory + "Supression List.xlsx");

            using (ExcelPackage package = new ExcelPackage(supFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["funds"];

                for (int columns = worksheet.Dimension.Start.Column; columns < worksheet.Dimension.End.Column + 1; columns++)
                {
                    if (worksheet.Cells[1, columns].Text.Contains("individual"))
                    {
                        for (int row = worksheet.Dimension.Start.Row + 1; row < worksheet.Dimension.End.Row; row++)
                        {
                            if (worksheet.Cells[row, columns].Value != null)
                            {
                                if (worksheet.Cells[row, columns].Value.ToString().Contains("@"))
                                {

                                    SupList.Add(supCellFormater(worksheet.Cells[row, columns].Value.ToString()));

                                }
                                else
                                {
                                    if (worksheet.Cells[row, columns - 1].Value != null)
                                    {
                                        SupList.Add(worksheet.Cells[row, columns - 1].Value.ToString());
                                    }
                                }
                            }
                        }
                        break;
                    }
                }
            }
        }

        public List<List<string>> FundList(string filePath)
        {
            FileInfo fundFile = new FileInfo(filePath);

            List<List<string>> fundList = new List<List<string>>();

            using (ExcelPackage package = new ExcelPackage(fundFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                for (int columns = worksheet.Dimension.Start.Column; columns < worksheet.Dimension.End.Column + 1; columns++)
                {
                    if (worksheet.Cells[1, columns].Text.Contains("mail") || worksheet.Cells[1, columns].Text.Contains("Mail") || worksheet.Cells[1, columns].Text.Contains("Email") || worksheet.Cells[1, columns].Text.Contains("email"))
                    {
                        for (int row = worksheet.Dimension.Start.Row + 1; row < worksheet.Dimension.End.Row; row++)
                        {
                            if (worksheet.Cells[row, columns].Value != null)
                            {
                                List<string> data = new List<string>
                                {
                                    row.ToString(),
                                    worksheet.Cells[row, columns].Value.ToString()
                                };
                                fundList.Add(data);
                            }
                        }
                    }
                }
            }
            return fundList;
        }


        private string supCellFormater(string cellValue)
        {
            return cellValue.Split('@')[1].Split('.')[0];
        }
    }
}
