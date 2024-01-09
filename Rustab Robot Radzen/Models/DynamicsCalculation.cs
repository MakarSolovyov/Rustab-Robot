using ASTRALib;
using Rustab_Robot_Radzen.Pages.Workpages;
using Rustab_Robot_Radzen.Shared;
using System.Data;

namespace Rustab_Robot_Radzen.Models
{
    /// <summary>
    /// Расчет предельного перетока классическим методом
    /// </summary>
    public class DynamicsCalculation : RastrSupplier
    {
        /// <summary>
        /// Функция для передачи сообщения в консоль
        /// </summary>
        /// <param name="FWDynamic">Экземпляр класса Rastr для расчета динамики</param>
        /// <param name="powerInSection">Переток в контролируемом сечении</param>
        private static void InvokeStopMessage(FWDynamic FWDynamic, double powerInSection)
        {
            MainLayout.console.Log(FWDynamic.ResultMessage);
            MainLayout.console.Log($"Динамическая устойчивость нарушается при перетоке в сечении, равном {powerInSection} МВт.");
        }

        /// <summary>
        /// Расчет предельного перетока классическим методом
        /// </summary>
        /// <param name="stepCounter">Счетчик шагов утяжеления</param>
        /// <param name="powerInSection">Мощность в сечении</param>
        /// <param name="researchingPlantGenerators">Исследуемая электростанция</param>
        /// <param name="scnFileNamesCopy">Имя файла сценария</param>
        /// <param name="scn">Сценарий возмущения</param>
        public static double? CalculateDynamics(string kprFilePath, int stepCounter, double powerInSection,
        List<int> researchingPlantGenerators, double modelingTime)
        {
            RastrSupplier.LoadFile(kprFilePath, Settings.shablonPaths[ShablonFileType.ControlledParameter]);

            FWDynamic FWDynamic = _rastr.FWDynamic();

            ITable table = _rastr.Tables.Item("com_dynamics");
            
            ICol columnItemSnap = table.Cols.Item("SnapPath");

            columnItemSnap.set_ZN(0, Settings.calcResultFolder);

            ICol columnItemTras = table.Cols.Item("Tras");
            columnItemTras.set_ZN(0, modelingTime);

            ICol columnItemHout = table.Cols.Item("Hout");
            columnItemHout.set_ZN(0, 0.01);

            ICol columnItemMaxResultFiles = table.Cols.Item("MaxResultFiles");
            columnItemMaxResultFiles.set_ZN(0, 1500);

            ICol columnItemCsv = table.Cols.Item("RealtimeCSV");
            columnItemCsv.set_ZN(0, true);            
            
            FWDynamic.Run();            
            FWDynamic.RunEMSmode(); // Расчёт переходного процесса

            double? powerInSectionPred = null;

            // Анализ результатов расчёта переходного процесса
            switch (FWDynamic.SyncLossCause)
            {
                case DFWSyncLossCause.SYNC_LOSS_NONE:
                    {
                        MainLayout.console.Log($"Потери синхронизма не выявлено.");
                        break;
                    }
                case DFWSyncLossCause.SYNC_LOSS_COA:
                    {
                        MainLayout.console.Log($"Выявлено превышение угла по сопротивлению генератора значения 180° в {FWDynamic.TimeReached} c.");
                        InvokeStopMessage(FWDynamic, powerInSection);

                        // Если мы уже сделали сколько-то шагов, можем вычислить Рпред
                        if (stepCounter != 0)
                        {
                            powerInSectionPred = PowerInSectionCounter(researchingPlantGenerators, powerInSection, stepCounter);
                        }
                        else
                        {
                            IfStabilityLossWithMinGen(powerInSection, stepCounter);
                        }

                        // Удаляем этот сценарий, потому что он больше не нужен - предел по ДУ найден
                        // scnFileNamesCopy.Remove(scn); 

                        break;
                    }
                case DFWSyncLossCause.SYNC_LOSS_BRANCHANGLE:
                    {
                        MainLayout.console.Log($"Выявлено превышение угла по ветви значения 180° в {FWDynamic.TimeReached} c.");
                        InvokeStopMessage(FWDynamic, powerInSection);

                        // Если мы уже сделали сколько-то шагов, можем вычислить Рпред
                        if (stepCounter != 0)
                        {
                            powerInSectionPred = PowerInSectionCounter(researchingPlantGenerators, powerInSection, stepCounter);
                        }
                        else
                        {
                            IfStabilityLossWithMinGen(powerInSection, stepCounter);
                        }

                        // Удаляем этот сценарий, потому что он больше не нужен - предел по ДУ найден
                        // scnFileNamesCopy.Remove(scn); 

                        break;
                    }
                case DFWSyncLossCause.SYNC_LOSS_OVERSPEED:
                    {
                        MainLayout.console.Log($"Выявлено превышение допустимой скорости вращения генератора в {FWDynamic.TimeReached} c.");
                        InvokeStopMessage(FWDynamic, powerInSection);

                        // Если мы уже сделали сколько-то шагов, можем вычислить Рпред
                        if (stepCounter != 0)
                        {
                            powerInSectionPred = PowerInSectionCounter(researchingPlantGenerators, powerInSection, stepCounter);
                        }
                        else
                        {
                            IfStabilityLossWithMinGen(powerInSection, stepCounter);
                        }

                        // Удаляем этот сценарий, потому что он больше не нужен - предел по ДУ найден
                        // scnFileNamesCopy.Remove(scn); 

                        break;
                    }
            }

            return powerInSectionPred;
        }

    }
}
