using ASTRALib;
using Rustab_Robot_Radzen.Models.DSI;
using Rustab_Robot_Radzen.Models.InfluentFactors;
using Rustab_Robot_Radzen.Pages.Workpages;
using Rustab_Robot_Radzen.Shared;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;

namespace Rustab_Robot_Radzen.Models
{
    public abstract class CalculationOrder : RastrSupplier
    {
        public static List<CSVDSIResult> StartCalculation(List<Scheme> schemeList, List<InfluentFactorBase> factorList,
           List<int> researchingPlantGenerators, int ResearchingSectionNumber, List<string> scnFileNames,
           CancellationToken cancellationToken, CalculationTask calculationTask, string scnInfoFilePath,
           string pPredInfoFilePath, string kprFileName, CalculationType calculationType)
        {
            string nameOfSection = GetStringValue("sechen", "ns", ResearchingSectionNumber, "name");
            MainLayout.console.Log($"Исследуемое сечение - '{nameOfSection}'.");

            List<CSVDSIResult> dsiList = new List<CSVDSIResult>();

            foreach (var scheme in schemeList)
            {
                MainLayout.console.Log($"Исследуемая схема - '{scheme.SchemeName}'.");
                MainLayout.console.Log($"Переток в исследуемом сечении {ResearchingSectionNumber} до утяжеления -" +
                    $" {Math.Round(GetValue("sechen", "ns", ResearchingSectionNumber, "psech"), 0)} МВт.");

                int maxIteration = 100;
                int iteration = 0;

                string shablonRst = Settings.shablonPaths[ShablonFileType.Scheme];
                LoadFile(scheme.SchemePath, shablonRst);

                List<string> scnList = new List<string>();
                for (var scnIndex = 0; scnIndex < scheme.ScnList.Count; scnIndex++)
                {
                    scnList.Add(scnFileNames[scheme.ScnList[scnIndex] - 1]);
                }

                List<double?> pPredList = new List<double?>();

                if (calculationType == CalculationType.Classic)
                {
                    pPredList = OneSchemeCalculation(factorList, maxIteration, researchingPlantGenerators, iteration,
                        ResearchingSectionNumber, scheme.SchemePath, scnList, cancellationToken, CalculationType.Classic, kprFileName);

                    //List<CSVDSIResult> dsiList = new List<CSVDSIResult>();
                    DSICalculation dsiCalc = new DSICalculation();

                    dsiList.AddRange(dsiCalc.CalculateDSI(scnInfoFilePath, pPredInfoFilePath, scheme.ScnList, pPredList));
                }
                //else if (calculationType == CalculationType.Ranking)
                //{
                //    pPredList = OneSchemeCalculation(factorList, maxIteration, researchingPlantGenerators, iteration,
                //        ResearchingSectionNumber, scheme.SchemePath, scnList, cancellationToken, CalculationType.Ranking, kprFileName);

                //    MainLayout.console.Log($"Конец расчета данных для определения показателей тяжести.");
                //    MainLayout.console.Log($"Запуск расчета показателей тяжести.");

                //    //List<CSVDSIResult> dsiList = new List<CSVDSIResult>();
                //    DSICalculation dsiCalc = new DSICalculation();

                //    dsiList = dsiCalc.CalculateDSI(scnInfoFilePath, pPredInfoFilePath, scheme.ScnList, pPredList);

                //    MainLayout.console.Log($"Показатели тяжести рассчитаны. Количество расчетных случаев: {dsiList.Count}.");

                //    dsiList = dsiCalc.RangeDSI(dsiList);

                //    List<string> scnFileNamesAfterRanking = new();
                //    if (calculationTask == CalculationTask.MaxPowerFlow)
                //    {
                //        MainLayout.console.Log($"Проведено ранжирование смоделированных возмущений. " +
                //        $"Наиболее тяжелое возмущение - сценарий № {dsiList[0].ScnNumber}.");
                //        MainLayout.console.Log($"Запуск расчета динамической устойчивости для сценария № {dsiList[0].ScnNumber}.");

                //        scnFileNamesAfterRanking.Add(scnFileNames[dsiList[0].ScnNumber - 1]);

                //        MainLayout.console.Log($"Переток в исследуемом сечении {ResearchingSectionNumber} до утяжеления -" +
                //        $" {Math.Round(GetValue("sechen", "ns", ResearchingSectionNumber, "psech"), 0)} МВт.");

                //        maxIteration = 100;
                //        iteration = 0;
                //        OneSchemeCalculation(factorList, maxIteration, researchingPlantGenerators, iteration,
                //            ResearchingSectionNumber, scheme.SchemePath, scnFileNamesAfterRanking, cancellationToken, CalculationType.Classic, kprFileName);

                //        // TODO: Придумать, какой вывод будет
                //    }
                //    else if (calculationTask == CalculationTask.ControlAction)
                //    {
                //        // TODO: Придумать, какой вывод будет
                //    }
                //}
            }

            return dsiList;
        }


        //public static void ClassicCalculation(List<Scheme> schemeList, List<InfluentFactorBase> factorList,
        //   List<int> researchingPlantGenerators, int ResearchingSectionNumber, List<string> scnFileNames,
        //   CancellationToken cancellationToken, CalculationTask calculationTask, string scnInfoFilePath, string pPredInfoFilePath, string kprFileName)
        //{
        //    string nameOfSection = GetStringValue("sechen", "ns", ResearchingSectionNumber, "name");
        //    MainLayout.console.Log($"Исследуемое сечение - '{nameOfSection}'.");

        //    foreach (var scheme in schemeList)
        //    {
        //        MainLayout.console.Log($"Исследуемая схема - '{scheme.SchemeName}'.");
        //        MainLayout.console.Log($"Переток в исследуемом сечении {ResearchingSectionNumber} до утяжеления -" +
        //            $" {Math.Round(GetValue("sechen", "ns", ResearchingSectionNumber, "psech"), 0)} МВт.");

        //        int maxIteration = 100;
        //        int iteration = 0;

        //        string shablonRst = Settings.shablonPaths[ShablonFileType.Scheme];
        //        LoadFile(scheme.SchemePath, shablonRst);

        //        List<string> scnList = new List<string>();
        //        for (var scnIndex = 0; scnIndex < scheme.ScnList.Count; scnIndex++)
        //        {
        //            scnList.Add(scnFileNames[scheme.ScnList[scnIndex] - 1]);
        //        }

        //        OneSchemeCalculation(factorList, maxIteration, researchingPlantGenerators, iteration,
        //            ResearchingSectionNumber, scheme.SchemePath, scnList, cancellationToken, CalculationType.Classic, kprFileName);
        //    }

        //    // TODO: Придумать, какой вывод будет у классического расчета
        //}

        //public static void RankingCalculation(List<Scheme> schemeList, List<InfluentFactorBase> factorList,
        //   List<int> researchingPlantGenerators, int ResearchingSectionNumber, List<string> scnFileNames,
        //   CancellationToken cancellationToken, CalculationTask calculationTask, string scnInfoFilePath, string pPredInfoFilePath, string kprFileName)
        //{
        //    string nameOfSection = GetStringValue("sechen", "ns", ResearchingSectionNumber, "name");
        //    MainLayout.console.Log($"Исследуемое сечение - '{nameOfSection}'.");

        //    foreach (var scheme in schemeList)
        //    {
        //        MainLayout.console.Log($"Исследуемая схема - '{scheme.SchemeName}'.");               
        //        MainLayout.console.Log($"Запуск расчета данных для определения показателей тяжести.");
        //        int maxIteration = 100;
        //        int iteration = 0;

        //        string shablonRst = Settings.shablonPaths[ShablonFileType.Scheme];
        //        LoadFile(scheme.SchemePath, shablonRst);

        //        List<string> scnList = new List<string>();
        //        for (var scnIndex = 0; scnIndex < scheme.ScnList.Count; scnIndex++)
        //        {
        //            scnList.Add(scnFileNames[scheme.ScnList[scnIndex] - 1]);
        //        }

        //        OneSchemeCalculation(factorList, maxIteration, researchingPlantGenerators, iteration,
        //            ResearchingSectionNumber, scheme.SchemePath, scnList, cancellationToken, CalculationType.Ranking, kprFileName);
                
        //        MainLayout.console.Log($"Конец расчета данных для определения показателей тяжести.");             
        //        MainLayout.console.Log($"Запуск расчета показателей тяжести.");

        //        List<CSVDSIResult> dsiList = new List<CSVDSIResult>();
        //        DSICalculation dsiCalc = new DSICalculation();

        //        dsiList = dsiCalc.CalculateDSI(scnInfoFilePath, pPredInfoFilePath, scheme.ScnList);

        //        MainLayout.console.Log($"Показатели тяжести рассчитаны. Количество расчетных случаев: {dsiList.Count}.");

        //        dsiList = dsiCalc.RangeDSI(dsiList);

        //        List<string> scnFileNamesAfterRanking = new();
        //        if (calculationTask == CalculationTask.MaxPowerFlow)
        //        {
        //            MainLayout.console.Log($"Проведено ранжирование смоделированных возмущений. " +
        //            $"Наиболее тяжелое возмущение - сценарий № {dsiList[0].ScnNumber}.");
        //            MainLayout.console.Log($"Запуск расчета динамической устойчивости для сценария № {dsiList[0].ScnNumber}.");

        //            scnFileNamesAfterRanking.Add(scnFileNames[dsiList[0].ScnNumber - 1]);

        //            MainLayout.console.Log($"Переток в исследуемом сечении {ResearchingSectionNumber} до утяжеления -" +
        //            $" {Math.Round(GetValue("sechen", "ns", ResearchingSectionNumber, "psech"), 0)} МВт.");

        //            maxIteration = 100;
        //            iteration = 0;
        //            OneSchemeCalculation(factorList, maxIteration, researchingPlantGenerators, iteration,
        //                ResearchingSectionNumber, scheme.SchemePath, scnFileNamesAfterRanking, cancellationToken, CalculationType.Classic, kprFileName);
        //        }
        //        else if (calculationTask == CalculationTask.ControlAction)
        //        {
        //            // TODO: Придумать, какой вывод будет
        //        }
        //    }

        //    // TODO: Придумать, какой вывод будет у расчета с ранжированием
        //}

        public static List<double?> OneSchemeCalculation(List<InfluentFactorBase> factorList, int maxIteration,
           List<int> researchingPlantGenerators, int iteration, int ResearchingSectionNumber, string rstFileName,
           List<string> scnFileNames, CancellationToken cancellationToken, CalculationType calculationType, string kprFileName)
        {
            // string nameOfSection = GetStringValue("sechen", "ns", ResearchingSectionNumber, "name");

            string shablonRst = Settings.shablonPaths[ShablonFileType.Scheme];
            string shablonScn = Settings.shablonPaths[ShablonFileType.Scenarios];
            List<string> scnFileNamesCopy = new List<string>();
            int stepCounter = 0;
            
            scnFileNamesCopy.AddRange(scnFileNames);

            List<double?> pPredList = new List<double?>();

            RastrRetCode kod, kd;
            if (_rastr.ut_Param[ParamUt.UT_FORM_P] == 0) //формировать описание КВ
            {
                #region Задание параметров утяжеления и инициализация таблиц
                _rastr.Tables.Item("ut_common").Cols.Item("tip").Z[0] = 0; //тип утяжеления - стандартный
                _rastr.ut_FormControl(); //формирует таблицу описаний контролируемых величин
                _rastr.ClearControl(); //инициализировать таблицу значений контролируемых величин

                kod = _rastr.step_ut("i"); //"i" – инициализировать значения параметров утяжеления (шаг в этом случае не выполняется)
                #endregion

                if (kod == 0)
                {
                    do
                    {
                        // Переток до следующего шага по траектории 
                        double powerInSection = Math.Round(GetValue("sechen", "ns", ResearchingSectionNumber, "psech"), 0);

                        if (scnFileNamesCopy.Count == 0) //если закончились сценарии для расчёта, то завершаем работу алгоритма
                        {
                            MainLayout.console.Log($"Расчёт переходных процессов по всем сценариям успешно завершён на {stepCounter} шаге утяжеления.");
                            
                            return pPredList;
                        }

                        foreach (string scn in scnFileNamesCopy.ToArray())
                        {
                            try
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                LoadFile(scn, shablonScn);

                                FileInfo info = new FileInfo(scn);
                                string name = info.Name;

                                MainLayout.console.Log($"Запущен расчёт по сценарию '{scn}'.");

                                double? pPredValue = null;

                                switch (calculationType)
                                {
                                    case CalculationType.Classic:
                                        {
                                            pPredValue = DynamicsCalculation.CalculateDynamics(kprFileName, stepCounter, powerInSection,
                                                researchingPlantGenerators, Settings.calcTimevalue);
                                            break;
                                        }
                                    case CalculationType.Ranking:
                                        {
                                            pPredValue = DynamicsCalculation.CalculateDynamics(kprFileName, stepCounter, powerInSection,
                                                researchingPlantGenerators, Settings.calcDSITimevalue);
                                            break;
                                        }
                                }

                                pPredList.Add(pPredValue);
                            }
                            catch (Exception ex)
                            {
                                MainLayout.console.Log($"Ошибка! {ex.Message}");
                                
                                return pPredList;
                            }
                        }

                        MainLayout.console.Log($"Расчёт переходных процессов на шаге {stepCounter} окончен.");
                        cancellationToken.ThrowIfCancellationRequested();

                        //проверка, не дошли ли генераторы до максимума своих регулировочных способностей
                        for (int j = 0; j < researchingPlantGenerators.Count; j++)
                        {
                            double powerOfGen = GetValue("Generator", "Node",  researchingPlantGenerators[j], "P");
                            double powerOfGenMax = GetValue("Generator", "Node", researchingPlantGenerators[j], "Pmax");

                            if (powerOfGen >= powerOfGenMax)
                            {
                                MainLayout.console.Log($"Расчёт окончен. Генераторы достигли предельной загрузки.");
                                MainLayout.console.Log($"Переток в исследуемом сечении {ResearchingSectionNumber} - " +
                                    $"{Math.Round(GetValue("sechen", "ns", ResearchingSectionNumber, "psech"), 0)} МВт.");
                                
                                return pPredList;
                            }
                        }

                        //еще одна проверка, перед тем, как шагнуть
                        if (scnFileNamesCopy.Count == 0) //если закончились сценарии для расчёта, то завершаем работу алгоритма
                        {
                            MainLayout.console.Log($"Расчёт переходных процессов по всем сценариям успешно завершён на {stepCounter} шаге утяжеления.");
                            return pPredList;
                        }

                        cancellationToken.ThrowIfCancellationRequested();

                        // шаг утяжеления
                        kd = _rastr.step_ut("z");
                        if (((kd == 0) && (_rastr.ut_Param[ParamUt.UT_ADD_P] == 0))
                            || _rastr.ut_Param[ParamUt.UT_TIP] == 1)
                        {
                            _rastr.AddControl(-1, ""); //Добавить строку в таблицу значений контролируемых величин
                            stepCounter++;
                            MainLayout.console.Log($"Выполняется {stepCounter} шаг утяжеления...");
                        }
                        
                        // Проверка находятся ли влияющие факторы в заданных пределах
                        foreach (var factor in factorList)
                        {
                            switch (factor)
                            {
                                case SectionFactor:
                                    {
                                        factor.CurrentValue = GetValue("sechen", "ns", factor.NumberFromRastr, "psech");

                                        if (!InfluentFactorBase.IsInDiapasone(factor))
                                        {
                                            MainLayout.console.Log($"Влияющий фактор {factor.FactorType} c номером " +
                                                $"{factor.NumberFromRastr} вышел за границу диапазона.");

                                            do
                                            {
                                                //factor.CurrentValue = GetValue("sechen", "ns", factor.NumberFromRastr, "psech");
                                                StepBack();

                                                FileInfo fileInfo = new FileInfo(rstFileName);
                                                rstFileName = $@"{fileInfo.DirectoryName}\{DateTime.Now.ToString().Replace(":",
                                                    "-")}{fileInfo.Extension}";
                                                SaveFile(rstFileName, shablonRst);
                                                LoadFile(rstFileName, shablonRst);

                                                cancellationToken.ThrowIfCancellationRequested();
                                                SectionFactor.CorrectTrajectory(factor);

                                                iteration += 1;

                                                if (iteration > maxIteration)
                                                {
                                                    MainLayout.console.Log($"Расчёт остановлен. Коррекция траектории утяжеления " +
                                                        $"по заданным исходным данным невозможна. Попробуйте ещё раз.");

                                                    return pPredList;
                                                }

                                                LoadFile(rstFileName, shablonRst);

                                                RastrRetCode kd2;
                                                // шаг утяжеления
                                                kd2 = _rastr.step_ut("z");
                                                if (((kd2 == 0) && (_rastr.ut_Param[ParamUt.UT_ADD_P] == 0))
                                                    || _rastr.ut_Param[ParamUt.UT_TIP] == 1)
                                                {
                                                    _rastr.AddControl(-1, ""); //Добавить строку в таблицу значений контролируемых величин
                                                }

                                                factor.CurrentValue = GetValue("sechen", "ns", factor.NumberFromRastr, "psech");
                                            }
                                            while (!InfluentFactorBase.IsInDiapasone(factor));
                                            //StepBack();

                                            MainLayout.console.Log($"Влияющий фактор {factor.FactorType} " +
                                                $"c номером {factor.NumberFromRastr} скорректирован и находится в заданном диапазоне.");
                                        }

                                        break;
                                    }
                                case VoltageFactor:
                                    {
                                        factor.CurrentValue = GetValue("node", "ny", factor.NumberFromRastr, "vras");

                                        if (!InfluentFactorBase.IsInDiapasone(factor))
                                        {
                                            MainLayout.console.Log($"Влияющий фактор {factor.FactorType} " +
                                                $"c номером {factor.NumberFromRastr} вышел за границу диапазона.");
                                            do
                                            {
                                                cancellationToken.ThrowIfCancellationRequested();
                                                VoltageFactor.CorrectVoltage(researchingPlantGenerators, factor);
                                                Regime();
                                                factor.CurrentValue = GetValue("node", "ny", factor.NumberFromRastr, "vras");
                                                iteration += 1;

                                                if (iteration > maxIteration)
                                                {
                                                    MainLayout.console.Log($"Расчёт остановлен. Коррекция траектории утяжеления " +
                                                        $"по заданным исходным данным невозможна. Попробуйте ещё раз.");
                                                    return pPredList;
                                                }
                                            }
                                            while (!InfluentFactorBase.IsInDiapasone(factor));

                                            MainLayout.console.Log($"Влияющий фактор {factor.FactorType} " +
                                                $"c номером {factor.NumberFromRastr} скорректирован и находится в заданном диапазоне.");
                                        }

                                        break;
                                    }
                            }
                        }

                        //если все факторы попали в диапазон, всё хорошо, шагаем дальше

                        //если генераторы дошли до максимума, останавливается расчёт
                        for (int j = 0; j < researchingPlantGenerators.Count; j++)
                        {
                            double powerOfGen = GetValue("Generator", "Node", researchingPlantGenerators[j], "P");
                            double powerOfGenMax = GetValue("Generator", "Node", researchingPlantGenerators[j], "Pmax");

                            if (powerOfGen >= powerOfGenMax)
                            {
                                MainLayout.console.Log($"Расчёт окончен. " +
                                    "Генераторы достигли предельной загрузки. Расчёт окончен.");
                                return pPredList;
                            }
                        }
                    }
                    while (kd == 0);
                }
            }

            // MainLayout.console.Log($"Превышено предельное число итераций!");
            MainLayout.console.Log($"Расчёт для рассматриваемой схемы завершён на {stepCounter} шаге утяжеления.");
            MainLayout.console.Log($"Величина перетока в исследуемом" +
                $" сечении {ResearchingSectionNumber} составляет" +
                $" {Math.Round(GetValue("sechen", "ns", ResearchingSectionNumber, "psech"), 0)} МВт.");

            return pPredList;
        }
    }
}
