namespace Rustab_Robot_Radzen.Models.InfluentFactors
{
    /// <summary>
    /// Влияющий фактор - Переток в сечении
    /// </summary>
    [Serializable]
    public class SectionFactor : InfluentFactorBase
    {
        /// <summary>
        /// Поле для реакции перетока в сечении на шаг по траектории
        /// </summary>
        private double _reaction;

        /// <summary>
        /// Тип фактора: Переток в сечении
        /// </summary>
        public override string FactorType => "Переток в сечении";

        /// <summary>
        /// Поле для списка генераторов, поддерживающих переток в этом сечении (влияющая станция)
        /// </summary>
        private List<int> _regulatingGeneratorsList = new List<int>();

        /// <summary>
        /// Список генераторов, поддерживающих переток в этом сечении (влияющая станция)
        /// </summary>
        public List<int> RegulatingGeneratorsList
        {
            get
            {
                return _regulatingGeneratorsList;
            }

            set
            {
                _regulatingGeneratorsList = value;
            }
        }

        /// <summary>
        /// Реакция перетока в сечении на шаг по траектории
        /// </summary>
        public double Reaction
        {
            get
            {
                return _reaction;
            }

            set
            {
                _reaction = value;
            }
        }

        /// <summary>
        /// Коррекция траектории утяжеления в процессе расчёта
        /// </summary>
        /// <param name="sectionFactor">Экземпляр класса InfluentFactorBase: влияющий фактор перетока в сечении</param>
        public static void CorrectTrajectory(InfluentFactorBase sectionFactor)
        {
            // Для каждого генератора данного сечения
            for (int i = 0; i < ((SectionFactor)sectionFactor).RegulatingGeneratorsList.Count; i++)
            {
                double initialIncrement = RastrSupplier.GetValue("ut_node", "ny", ((SectionFactor)sectionFactor).RegulatingGeneratorsList[i], "pg");
                double addIncrement = 0;
                
                if (sectionFactor.CurrentValue < sectionFactor.MinValue)
                {
                    addIncrement = (sectionFactor.MaxValue - sectionFactor.MinValue) / (((SectionFactor)sectionFactor).RegulatingGeneratorsList.Count);
                }
                else if (sectionFactor.CurrentValue > sectionFactor.MaxValue)
                {
                    addIncrement = -(sectionFactor.MaxValue - sectionFactor.MinValue) / (((SectionFactor)sectionFactor).RegulatingGeneratorsList.Count);
                }

                // Приращение генерации пересчитанное
                RastrSupplier.SetValue("ut_node", "ny", ((SectionFactor)sectionFactor).RegulatingGeneratorsList[i], "pg", initialIncrement + addIncrement);
            }
        }

    }
}
