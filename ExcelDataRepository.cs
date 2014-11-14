using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Reflection;

namespace OkCupidAutoBot
{
    class ExcelDataRepository
    {
        public void SaveDetails(List<ProfileDetails> listOfProfiles, FileInfo outputPath)
        {
            //create the excel file and header rows if file does not exist
            if (!Helpers.DoesSheetExist(outputPath, "GirlsDetails"))
                CreateDetailsHeader(outputPath);

            using(ExcelPackage package = new ExcelPackage(outputPath))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["GirlsDetails"];

                //check for the last row in order to append to it
                int lastRow = worksheet.Dimension.End.Row;
                int idValue = default(int);
                bool isNumber = int.TryParse(worksheet.Cells[lastRow, 1].Value.ToString(), out idValue);

                if (!isNumber) //if the value is not a number, seed it with 1
                    idValue = 1; //seed
                else
                    idValue++;

                lastRow++; //append new row

                foreach (ProfileDetails profile in listOfProfiles)
                {
                    int startCol = 1;
                    worksheet.Cells[lastRow, startCol].Value = idValue;

                    Type pType = typeof(ProfileDetails);
                    foreach (PropertyInfo pInfo in pType.GetProperties())
                    {
                        startCol++;
                        worksheet.Cells[lastRow, startCol].Value = pInfo.GetValue(profile, null);
                    }

                    lastRow++; //increment 1 row
                    idValue++; //increment id
                }

                //format the sheet
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                worksheet.View.ShowGridLines = false;
                worksheet.View.FreezePanes(2, 1);

                //finally, save the sheet
                package.Save();
            }
        }

        public void CreateDetailsHeader(FileInfo outputPath)
        {
            using (ExcelPackage package = new ExcelPackage(outputPath))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("GirlsDetails");

                //headers
                int startCol = 1;
                worksheet.Cells[1, startCol].Value = "id";
                Type type = typeof(ProfileDetails);
                foreach (PropertyInfo pInfo in type.GetProperties())
                {
                    startCol++;
                    worksheet.Cells[1, startCol].Value = pInfo.Name;
                }

                string lastCol = GetColumnName(startCol);
                string headerRange = "A1:" + lastCol + "1";

                worksheet.Cells[headerRange].Style.Font.Bold = true; //set header row to bold

                //id col
                worksheet.Column(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                worksheet.Column(1).Style.Numberformat.Format = "0";

                //username col
                worksheet.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                //age col
                worksheet.Column(4).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                worksheet.Column(4).Style.Numberformat.Format = "0";

                //match percentage col
                worksheet.Column(6).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                worksheet.Column(6).Style.Numberformat.Format = "0";
                //enemy percentage col
                worksheet.Column(7).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                worksheet.Column(7).Style.Numberformat.Format = "0";

                //lastonline col
                worksheet.Column(8).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Column(8).Style.Numberformat.Format = "yyyy-MM-dd";

                //created date col
                worksheet.Column(28).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Column(28).Style.Numberformat.Format = "yyyy-MM-dd hh:mm";

                //format the sheet
                worksheet.View.ShowGridLines = false;
                worksheet.View.FreezePanes(2, 1);

                package.Save();
            }
        }

        public void CreateEssaysHeader(FileInfo outputPath)
        {
            using (ExcelPackage package = new ExcelPackage(outputPath))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("GirlsEssays");

                //create header rows
                int start = 1;
                worksheet.Cells[1, start].Value = "id";

                Type type = typeof(Essays);
                foreach (PropertyInfo pInfo in type.GetProperties())
                {
                    start++;
                    worksheet.Cells[1, start].Value = pInfo.Name;
                }

                string lastCol = GetColumnName(start);
                string headerRange = "A1:" + lastCol + "1";

                //created date col
                worksheet.Column(15).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Column(15).Style.Numberformat.Format = "yyyy-MM-dd hh:mm";

                worksheet.Cells[headerRange].Style.Font.Bold = true;

                package.Save();
            }
        }

        public void SaveEssays(List<Essays> essays, FileInfo outputPath)
        {
            if (!Helpers.DoesSheetExist(outputPath, "GirlsEssays"))
                CreateEssaysHeader(outputPath);

            using (ExcelPackage package = new ExcelPackage(outputPath))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["GirlsEssays"];

                //append
                int lastRow = worksheet.Dimension.End.Row;
                int idValue = default(int);
                bool isNumber = int.TryParse(worksheet.Cells[lastRow, 1].Value.ToString(), out idValue);

                if (!isNumber) //if the value is not a number, seed it with 1
                    idValue = 1; //seed
                else
                    idValue++;

                lastRow++; //append new row

                foreach (Essays essay in essays)
                {
                    int startCol = 1;
                    worksheet.Cells[lastRow, startCol].Value = idValue;

                    Type essayType = typeof(Essays);
                    foreach (PropertyInfo essayProp in essayType.GetProperties())
                    {
                        startCol++;
                        worksheet.Cells[lastRow, startCol].Value = essayProp.GetValue(essay, null);
                    }

                    //increment row and id value
                    lastRow++;
                    idValue++;
                }

                worksheet.View.ShowGridLines = false;

                //finally, save the sheet
                package.Save();
            }
        }

        public string GetColumnName(int columnNum)
        {
            int d, m;
            string name = default(string);

            d = columnNum;

            while (d > 0)
            {
                m = (d - 1) % 26;
                name = ((char)(65 + m)) + name;
                d = (d - m) / 26;
            }

            return name;
        }
    }
}
