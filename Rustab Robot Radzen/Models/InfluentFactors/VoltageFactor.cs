namespace Rustab_Robot_Radzen.Models.InfluentFactors
{
    /// <summary>
    /// Влияющий фактор - Напряжение на шинах исследуемой станции
    /// Регулируется генераторами исследуемой станции (изменение V_зд)
    /// </summary>
    [Serializable]
    public class VoltageFactor : InfluentFactorBase
    {
        /// <summary>
        /// Тип фактора: Напряжение
        /// </summary>
        public override string FactorType => "Напряжение в узле";

        /// <summary>
        /// Метод для коррекции V_зд у генераторов для поддержания напряжения-ВФ
        /// </summary>
        /// <param name="listOfGenerators">Список рассматриваемых генераторов</param>
        /// <param name="voltageFactor">Экземпляр класса InfluentFactorBase: влияющий фактор напряжения в узле</param>
        public static void CorrectVoltage(List<int> listOfGenerators, InfluentFactorBase voltageFactor)
        {
            for (int i = 0; i < listOfGenerators.Count; i++)
            {
                double vzdOfGen = RastrSupplier.GetValue("node", "ny", listOfGenerators[i], "vzd");
                double uhomOfGen = RastrSupplier.GetValue("node", "ny", listOfGenerators[i], "uhom");

                if (voltageFactor.CurrentValue < voltageFactor.MinValue)
                {
                    if (vzdOfGen < setMaxValueForVoltage(uhomOfGen))
                    {
                        RastrSupplier.SetValue("node", "ny", listOfGenerators[i], "vzd", vzdOfGen + 0.1);
                    }
                    else
                    {
                        RastrSupplier.SetValue("node", "ny", listOfGenerators[i], "vzd", setMaxValueForVoltage(uhomOfGen));
                    }
                }
                else if (voltageFactor.CurrentValue > voltageFactor.MaxValue)
                {
                    if (vzdOfGen > setMinValueForVoltage(uhomOfGen))
                    {
                        RastrSupplier.SetValue("node", "ny", listOfGenerators[i], "vzd", vzdOfGen - 0.1);
                    }
                    else
                    {
                        RastrSupplier.SetValue("node", "ny", listOfGenerators[i], "vzd", setMinValueForVoltage(uhomOfGen));
                    }
                }
            }
        }

        /// <summary>
        /// Максильмано возможное значение V_зд генератора
        /// </summary>
        /// <param name="vzd">Значение V_зд генератора</param>
        /// <returns>Максильмано возможное значение</returns>
        private static double setMaxValueForVoltage(double vzd)
        {
            return vzd * 1.05;
        }

        /// <summary>
        /// Минимально возможное значение V_зд генератора
        /// </summary>
        /// <param name="vzd">Значение V_зд генератора</param>
        /// <returns>Минимально возможное значение</returns>
        private static double setMinValueForVoltage(double vzd)
        {
            return vzd * 0.95;
        }
    }
}
