using System.Xml.Serialization;

namespace Rustab_Robot_Radzen.Models.InfluentFactors
{
    /// <summary>
    /// Базовый класс для влияющих факторов
    /// </summary>
    [Serializable]
    [XmlInclude(typeof(SectionFactor))]
    [XmlInclude(typeof(VoltageFactor))]
    public abstract class InfluentFactorBase
    {
        /// <summary>
        /// Нижняя граница для фактора
        /// </summary>
        private double _minValue;

        /// <summary>
        /// Верхняя граница для фактора
        /// </summary>
        private double _maxValue;

        /// <summary>
        /// Текущее значение фактора
        /// </summary>
        private double _currentValue;

        /// <summary>
        /// Идентификатор (номер) из RastrWin
        /// </summary>
        protected int _numberFromRastr;

        /// <summary>
        /// Тип влияющего фактора
        /// </summary>
        public abstract string FactorType { get; }

        /// <summary>
        /// Идентификатор (номер) из RastrWin
        /// </summary>
        public int NumberFromRastr
        {
            get
            {
                return _numberFromRastr;
            }
            set
            {
                _numberFromRastr = value;
            }
        }

        /// <summary>
        /// Минимальное значение для фактора
        /// </summary>
        public double MinValue
        {
            get
            {
                return _minValue;
            }
            set
            {
                _minValue = Math.Round(value, 0);
            }
        }

        /// <summary>
        /// Максимальное значение для фактора
        /// </summary>
        public double MaxValue
        {
            get
            {
                return _maxValue;
            }
            set
            {
                _maxValue = Math.Round(value, 0);
            }
        }

        /// <summary>
        /// Текущее значение фактора
        /// </summary>
        public double CurrentValue
        {
            get
            {
                return _currentValue;
            }
            set
            {
                _currentValue = Math.Round(value, 0);
            }
        }

        /// <summary>
        /// Проверка, входят ли факторы в диапазон (границы диапазона сравниваются с текущим значением)
        /// </summary>
        /// <param name="influentFactor">Экземпляр класса InfluentFactorBase</param>
        /// <returns>Логическое значение успешности проверки</returns>
        public static bool IsInDiapasone(InfluentFactorBase influentFactor)
        {
            if (influentFactor.CurrentValue > influentFactor.MaxValue || influentFactor.CurrentValue < influentFactor.MinValue)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Проверка, корректны ли границы диапазона, которые ввёл пользователь
        /// </summary>
        /// <param name="minValue">Минимальное значение для фактора</param>
        /// <param name="maxValue">Максимальное значение для фактора</param>
        /// <returns>Логическое значение успешности проверки</returns>
        public static bool IsMinMaxCorrect(double minValue, double maxValue)
        {
            if (minValue > maxValue || maxValue == minValue)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
    }
}
