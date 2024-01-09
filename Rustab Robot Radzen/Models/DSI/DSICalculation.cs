using CsvHelper;
using CsvHelper.Configuration;
using Rustab_Robot_Radzen.Pages.Workpages;
using Rustab_Robot_Radzen.Shared;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Rustab_Robot_Radzen.Models.DSI
{
    public class DSICalculation
    {
        //// Переместить список в файл формирования схем и передать сюда как параметр
        //List<int> scnNumbersList = new List<int>()
        //    { 13, 14, 17, 18, 21, 27, 28, 30, 31, 33, 34, 37, 38, 41, 42, 45, 46, 49, 50, 52, 53, 55, 56, 58, 59 }; // 21.1 схема

        public List<CSVDSIResult> CalculateDSI(string scnInfoFilePath, string pPredInfoFilePath, List<int> scnNumbersList, List<double?> pPredList)
        {
            Excel.Application scnInfoFile = new Excel.Application();
            Excel.Application pPredInfoFile = new Excel.Application();

            Excel.Workbook scnInfoWorkbook = scnInfoFile.Workbooks.Open(scnInfoFilePath);
            Excel.Worksheet scnInfoWorksheet = scnInfoWorkbook.Worksheets[1];
            Excel.Workbook pPredInfoWorkbook = pPredInfoFile.Workbooks.Open(pPredInfoFilePath);
            Excel.Worksheet pPredInfoWorksheet = pPredInfoWorkbook.Worksheets[1];

            List<CSVDSIResult> dsiList = new List<CSVDSIResult>();

            var tpCounter = 0;
            var scnCounter = 0;

            foreach (string resultFileName in GetFileNames(Settings.calcResultFolder))
            {
                var data = ReadCSVFile(resultFileName);

                var KZIndex = 0;
                var KZTime = 0.5;
                var counter = 0;
                var isStabilityLost = false;

                var initDelta = data[0].DeltaBoGESGen1;

                for (int i = 0; i < data.Count; i++)
                {
                    if ((data[i].Time < KZTime) && (counter == 0))
                    {
                        KZIndex++;
                    }
                    else
                    {
                        counter = 1;
                    }

                    var deltaValues = new Dictionary<string, double>()
                    {
                        ["Delta_Haranor_G3"] = data[i].DeltaHaranorGen3,
                        ["Delta_Belovo_G1"] = data[i].DeltaBelovoGen1,
                        ["Delta_Kras_G5"] = data[i].DeltaKrasGen5,
                        ["Delta_Bratsk_G4"] = data[i].DeltaBratskGen4,
                        ["Delta_Irkutsk_G2"] = data[i].DeltaIrkutskGen2,
                        ["Delta_Gusi_G5"] = data[i].DeltaGusiGen5,
                        ["Delta_Chita_G4"] = data[i].DeltaChitaGen4
                    };

                    foreach (var delta in deltaValues)
                    {                       
                        var delta5001 = Math.Abs(data[i].DeltaBoGESGen1 - delta.Value);
                        var delta5002 = Math.Abs(data[i].DeltaBoGESGen3 - delta.Value);
                        if ((delta5001 > 180) || (delta5002 > 180))
                        {
                            isStabilityLost = true;
                        }
                    }
                }

                var voltageDrop = data[KZIndex].UBoGES - data[KZIndex + 1].UBoGES;
                var genPowerDrop = data[KZIndex].PgBoGESGen1 - data[KZIndex + 1].PgBoGESGen1;
                if (genPowerDrop == 0)
                {
                    genPowerDrop = data[KZIndex].PgBoGESGen3 - data[KZIndex + 1].PgBoGESGen3;
                }

                var kzType = "";
                var kzLocation = "";
                double kzElimTime = 0;

                kzType = Convert.ToString(scnInfoWorksheet.Cells[scnNumbersList[scnCounter] + 1, 2].Value);
                kzLocation = Convert.ToString(scnInfoWorksheet.Cells[scnNumbersList[scnCounter] + 1, 3].Value);

                var subString = " вблизи";
                int indexOfChar = kzLocation.IndexOf(subString);
                var line = kzLocation.Substring(0, indexOfChar);

                kzElimTime = Convert.ToDouble(scnInfoWorksheet.Cells[scnNumbersList[scnCounter] + 1, 4].Value);
                

                var scheme = "21.1. Ремонт КВЛ 500 кВ Богучанская ГЭС - Ангара № 1 и КВЛ 500 кВ Богучанская ГЭС - Ангара № 2";
                var genNumber = 6;

                var genPower = Math.Round(data[0].PgBoGESGen1);
                if (genPower == 0)
                {
                    genPower = Math.Round(data[0].PgBoGESGen3);
                }

                double pPredPAR = 0;
                double maxPowerFlow = 0;

                for (int i = 2; i <= pPredInfoWorksheet.UsedRange.Rows.Count; i++)
                {
                    if ((Convert.ToString(pPredInfoWorksheet.Cells[i, 2].Value) == scheme) && (Convert.ToString(pPredInfoWorksheet.Cells[i, 3].Value) == line))
                    {
                        pPredPAR = Convert.ToDouble(pPredInfoWorksheet.Cells[i, 4].Value);
                        maxPowerFlow = Convert.ToDouble(pPredInfoWorksheet.Cells[i, 7].Value);
                    }
                }

                dsiList.Add(new CSVDSIResult
                {
                    ID = tpCounter + 1,
                    ScnNumber = scnNumbersList[scnCounter],
                    Scheme = scheme,
                    GenNumber = genNumber,
                    GenPower = genPower,
                    KZType = kzType,
                    KZLocation = kzLocation,
                    KZElimTime = kzElimTime,
                    VoltageDrop = voltageDrop,
                    GenPowerDrop = genPowerDrop,
                    PPredPAR = pPredPAR,
                    InitDelta = initDelta,
                    MaxPowerFlow = maxPowerFlow,
                    IsStabilityLost = isStabilityLost,
                    PPred = pPredList[tpCounter]
                });

                scnCounter++;
                if (scnCounter == scnNumbersList.Count)
                {
                    scnCounter = 0;
                }

                tpCounter++;
            }

            Marshal.ReleaseComObject(scnInfoWorksheet);
            Marshal.ReleaseComObject(pPredInfoWorksheet);
            scnInfoWorkbook.Close();
            pPredInfoWorkbook.Close();
            Marshal.ReleaseComObject(scnInfoWorkbook);
            Marshal.ReleaseComObject(pPredInfoWorkbook);
            scnInfoFile.Quit();
            pPredInfoFile.Quit();
            Marshal.ReleaseComObject(scnInfoFile);
            Marshal.ReleaseComObject(pPredInfoFile);

            return dsiList;
        }

        public List<CSVDSIResult> RangeDSI(List<CSVDSIResult> dsiList)
        {
            var maxKZElimTime = dsiList.Max(CSVDSIResult => CSVDSIResult.KZElimTime);
            var maxVoltageDrop = dsiList.Max(CSVDSIResult => CSVDSIResult.VoltageDrop);
            var maxGenPowerDrop = dsiList.Max(CSVDSIResult => CSVDSIResult.GenPowerDrop);
            var maxPPredPAR = dsiList.Max(CSVDSIResult => CSVDSIResult.PPredPAR) + 1;
            var minPPredPAR = dsiList.Min(CSVDSIResult => CSVDSIResult.PPredPAR);
            var maxInitDelta = dsiList.Max(CSVDSIResult => CSVDSIResult.InitDelta);

            foreach (var calcResult in dsiList)
            {
                var relKZElimTime = calcResult.KZElimTime / maxKZElimTime;
                var relVoltageDrop = calcResult.VoltageDrop / maxVoltageDrop;
                var relGenPowerDrop = calcResult.GenPowerDrop / maxGenPowerDrop;
                var relPPredPAR = (maxPPredPAR - calcResult.PPredPAR) / (maxPPredPAR - minPPredPAR);
                var relInitDelta = (180 - maxInitDelta) / (180 - calcResult.InitDelta);

                var complexDSI = relKZElimTime * relVoltageDrop * relGenPowerDrop * relPPredPAR * relInitDelta;
                calcResult.ComplexDSI = complexDSI;
            }

            var sortedDSIList = dsiList.OrderByDescending(dsi => dsi.ComplexDSI).ToList();

            return sortedDSIList;
        }

        private IEnumerable<string> GetFileNames(string DirName)
        {
            DirectoryInfo info = new DirectoryInfo(DirName);
            FileInfo[] finfo = info.GetFiles("*.csv");
            return finfo.Select(f => f.FullName);
        }

        public void WriteCSVFile(List<CSVDSIResult> dsiList)
        {
            string pathToSave = Path.Combine(Settings.rankResultFolder, "rankResultFile.csv");

            using (var writer = new StreamWriter(Settings.rankResultFolder, false, Encoding.UTF8))
            {
                var csvConfig = new CsvConfiguration(CultureInfo.GetCultureInfo("ru-RU"))
                {
                    Delimiter = ";"
                };

                using (var csv = new CsvWriter(writer, csvConfig))
                {
                    csv.WriteRecords(dsiList);
                }
            }  
        }

        private List<CSVCalcResult> ReadCSVFile(string filePath)
        {
            List<CSVCalcResult> resultInfo = new List<CSVCalcResult>();

            try
            {
                using (var writer = new StreamReader(filePath))
                {
                    var csvConfig = new CsvConfiguration(CultureInfo.GetCultureInfo("ru-RU"))
                    {
                        Delimiter = ";",
                        Encoding = Encoding.UTF8
                    };

                    using (var csv = new CsvReader(writer, csvConfig))
                    {
                        var record = new CSVCalcResult();
                        var records = csv.EnumerateRecords(record);

                        foreach (var r in records)
                        {
                            resultInfo.Add(new CSVCalcResult
                            {
                                Time = r.Time,
                                UBoGES = r.UBoGES,
                                DeltaBoGESGen1 = r.DeltaBoGESGen1,
                                DeltaBoGESGen3 = r.DeltaBoGESGen3,
                                PgBoGESGen1 = r.PgBoGESGen1,
                                PgBoGESGen3 = r.PgBoGESGen3,
                                DeltaHaranorGen3 = r.DeltaHaranorGen3,
                                DeltaBelovoGen1 = r.DeltaBelovoGen1,
                                DeltaKrasGen5 = r.DeltaKrasGen5,
                                DeltaSSHGESGen9 = r.DeltaSSHGESGen9,
                                DeltaBratskGen4 = r.DeltaBratskGen4,
                                DeltaIrkutskGen2 = r.DeltaIrkutskGen2,
                                DeltaGusiGen5 = r.DeltaGusiGen5,
                                DeltaChitaGen4 = r.DeltaChitaGen4
                            });
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                throw exception;            
            }

            return resultInfo;
        }
    }
}
