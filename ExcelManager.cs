using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParse
{
    public class ExcelManager
    {
        public static string brandName;
        public static string modelName;
        public static void WriteExcel(string fileName)
        {
           var fileInfo = new FileInfo(fileName);
            List<string> zeroProduct = new List<string>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(fileName))
            {
                using (ExcelWorksheet workSheet = package.Workbook.Worksheets[0])
                {
              
                    ModelPart[] modelParts = new ModelPart[workSheet.Dimension.End.Row - 1];
                   
                        int modelPartCount = 0;
                    var excelLit = workSheet.Cells[workSheet.Dimension.Start.Row + 1, 1, workSheet.Dimension.End.Row, 6].ToList();
                    foreach (var item in excelLit)
                    {
                        Console.WriteLine(item);
                    }
                    var excelList = workSheet?.Cells?[workSheet.Dimension.Start.Row+1 , 1, workSheet.Dimension.End.Row, 6]?.ToList();
                    Console.WriteLine(excelList);
                    for (int i = 0; i < excelList.Count-4;i+=4)
                    {
                        string str = excelList[i].Text;
                        if (IsBrandName(excelList[i]))
                        {
                            brandName = excelList[i].Text;
                           
                            continue;
                            
                        }
                        if (IsModelName(excelList[i]))
                        {
                            modelName = excelList[i].Text;
                           
                            continue;
                            
                        }
                        modelParts[modelPartCount] = new ModelPart();
                        modelParts[modelPartCount].BodyParts_Number = excelList[i].Text;
                        Console.WriteLine($"{modelParts[modelPartCount].BodyParts_Number}");
                        modelParts[modelPartCount].OEM_Number = excelList[i + 1].Text;
                        modelParts[modelPartCount].Year = excelList[i + 2].Text;
                        modelParts[modelPartCount].Description = excelList[i+3].Text;
                        modelParts[modelPartCount].Catalog = brandName;
                        modelParts[modelPartCount].Model = modelName;
                        modelPartCount++;
                        
                    }
                   
                }
            }
        }
        public static bool IsBrandName( OfficeOpenXml.ExcelRangeBase element)
        {
            if (element.Style.Fill.BackgroundColor.Indexed==22)
            {
                return true;
            }
            else return false;
        }
        public static bool IsModelName(OfficeOpenXml.ExcelRangeBase element)
        {
            if (element.Style.Fill.BackgroundColor.Indexed!=22 && element.Style.Font.Bold==true)
            {
                return true;
            }
            else return false;
        }

        public static void ReadExcel()
        {
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(@"C:/Users/glekli/source/repos/AdminServices/StaffService/CheckingTheCounterpartyService/Autopiter.CheckingApi/ExcelParse/ExcelParse/europa.xlsx", FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
       
            ISheet sheet = hssfwb.GetSheet("Лист1");
            List<ICell> list = new();
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                for (int cell = 0; cell < 4; cell++)
                {
                   
                    list.Add(sheet.GetRow(row).GetCell(cell));
                 
                    //Console.WriteLine(sheet.GetRow(row).GetCell(cell).StringCellValue);
                }
               
            }
        }
    }
}
