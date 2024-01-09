using CsvHelper.Configuration.Attributes;
using Rustab_Robot_Radzen.Pages.Workpages;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Rustab_Robot_Radzen.Models
{
    public class SchemeCreation
    {
        public List<Scheme> CreateSchemeList(string excelFilePath, string rstFilePath)
        {
            Excel.Application schemeInfoFile = new Excel.Application();

            Excel.Workbook schemeInfoWorkbook = schemeInfoFile.Workbooks.Open(excelFilePath);
            Excel.Worksheet schemeInfoWorksheet = schemeInfoWorkbook.Worksheets[1];
            
            List<Scheme> schemeList = new List<Scheme>();
            List<int> distinctSchemeNumbers = new List<int>();

            for (int row = 2; row <= schemeInfoWorksheet.UsedRange.Rows.Count - 1; row++)
            {
                var schemeNumber = Convert.ToInt32(schemeInfoWorksheet.Cells[row, 1].Value);
                var schemeName = Convert.ToString(schemeInfoWorksheet.Cells[row, 2].Value);
                var schemeGenNumber = Convert.ToInt32(schemeInfoWorksheet.Cells[row, 3].Value);

                var ipNumber = Convert.ToInt32(schemeInfoWorksheet.Cells[row, 4].Value);
                var iqNumber = Convert.ToInt32(schemeInfoWorksheet.Cells[row, 5].Value);
                var npNumber = Convert.ToInt32(schemeInfoWorksheet.Cells[row, 6].Value);

                RastrSupplier.ChangeVetvState(ipNumber, iqNumber, npNumber, true);

                if (schemeNumber != Convert.ToInt32(schemeInfoWorksheet.Cells[row + 1, 1].Value))
                {
                    var genNumbers = Convert.ToString(schemeInfoWorksheet.Cells[row, 7].Value);
                    string[] genNumbersArray = genNumbers.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                    var allGenNumbers = Convert.ToString(schemeInfoWorksheet.Cells[2, 8].Value);
                    string[] allGenNumbersArray = allGenNumbers.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                    var scnNumbers = Convert.ToString(schemeInfoWorksheet.Cells[row, 9].Value);
                    string[] scnNumbersArray = scnNumbers.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                    int[] scnNumbersArrayInt = scnNumbersArray.Select(x => int.Parse(x)).ToArray();

                    foreach (var gen in allGenNumbersArray)
                    {
                        if (genNumbersArray.Contains(gen))
                        {
                            var genID = Convert.ToInt32(gen);
                            RastrSupplier.SetBoolValue("node", "ny", genID, "sta", false);
                        }
                        else
                        {
                            var genID = Convert.ToInt32(gen);
                            RastrSupplier.SetBoolValue("node", "ny", genID, "sta", true);
                        }
                    }

                    string fileName = $"{schemeGenNumber} ген {schemeName}.rst";
                    string pathToSave = Path.Combine(Environment.CurrentDirectory, "CreatedFiles", "Schemes", fileName);

                    if (!distinctSchemeNumbers.Contains(schemeNumber))
                    {
                        distinctSchemeNumbers.Add(schemeNumber);

                        schemeList.Add(new Scheme
                        {
                            SchemeNumber = schemeNumber,
                            SchemeName = schemeName,
                            SchemeGenNumber = schemeGenNumber,
                            SchemePath = pathToSave,
                            ScnList = scnNumbersArrayInt.ToList()
                        });
                    }

                    RastrSupplier.Regime();
                    RastrSupplier.SaveFile(pathToSave, Settings.shablonPaths[ShablonFileType.Scheme]);

                    RastrSupplier.LoadFile(rstFilePath, Settings.shablonPaths[ShablonFileType.Scheme]);
                }  
            }

            Marshal.ReleaseComObject(schemeInfoWorksheet);
            schemeInfoWorkbook.Close();
            Marshal.ReleaseComObject(schemeInfoWorkbook);
            schemeInfoFile.Quit();
            Marshal.ReleaseComObject(schemeInfoFile);

            return schemeList;
        }
    }

    public class Scheme
    {
        [Name("№ схемы")]
        public int SchemeNumber { get; set; }
        [Name("Название схемы")]
        public string SchemeName { get; set; }
        [Name("Количество вкл. генераторов")]
        public int SchemeGenNumber { get; set; }
        [Name("Путь до файла")]
        public string SchemePath {  get; set; }
        [Name("Сценарии для расчета")]
        public List<int> ScnList { get; set; }
    }
}
