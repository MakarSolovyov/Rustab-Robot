using ASTRALib;
using Rustab_Robot_Radzen.Models.InfluentFactors;
using Rustab_Robot_Radzen.Pages.Workpages;
using Rustab_Robot_Radzen.Shared;
using System.ComponentModel;
using System.Data;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Rustab_Robot_Radzen.Models
{
    /// <summary>
    /// Класс, осуществляющий взаимодействие с RastrWin3
    /// </summary>
    public abstract class RastrSupplier
    {
        /// <summary>
        /// Экземпляр класса Rastr
        /// </summary>
        public static Rastr _rastr = new Rastr();

        /// <summary>
        /// Загрузка файла в рабочую область
        /// </summary>
        /// <param name="filePath">Путь файла для загрузки</param>
        /// <param name="shablon">Шаблон нужного файла</param>
        public static void LoadFile(string filePath, string shablon)
        {
            _rastr.Load(RG_KOD.RG_REPL, filePath, shablon);
        }

        /// <summary>
        /// Сохранение файла с заданием имени
        /// </summary>
        /// <param name="fileName">Задаваемое имя файла</param>
        /// <param name="shablon">Шаблон нужного файла</param>
        public static void SaveFile(string fileName, string shablon)
        {
            // Если нет лицензии RastrWin3, то файл не сохранится
            try
            {
                _rastr.Save(fileName, shablon);
            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message);
            }
        }

        /// <summary>
        /// Создание файла с помощью загрузки шаблона
        /// </summary>
        /// <param name="shablon">Шаблон нужного файла</param>
        public static void CreateFile(string shablon)
        {
            _rastr.NewFile(shablon);
        }

        /// <summary>
        /// Узнать индекс из таблицы по номеру 
        /// </summary>
        /// <param name="tableName">Название таблицы, в которой ведется поиск</param>
        /// <param name="parameterName">Название параметра, по которому ищется индекс</param>
        /// <param name="number">Номер узла (может быть так же номером сечения)</param>
        public static int GetIndexByNumber(string tableName, string parameterName, int number)
        {
            ITable table = _rastr.Tables.Item(tableName);
            ICol columnItem = table.Cols.Item(parameterName);

            for (int index = 0; index < table.Count; index++)
            {
                if (columnItem.get_ZN(index) == number)
                {
                    return index;
                }
            }
            throw new Exception($"Узел/сечение с номером {number} не найден(о).");
        }

        /// <summary>
        /// Задать булевое значение
        /// </summary>
        /// <param name="tableName">Название таблицы, в которой ведется поиск</param>
        /// <param name="parameterName">Название параметра, по которому ищется индекс</param>
        /// <param name="number">Номер узла (может быть так же номером сечения)</param>
        /// <param name="chosenParameter">Любой параметр, значение из ячеек которого нужно получить
        /// (например, модуль текущего напряжения в узле)</param>
        /// <param name="value">Записываемое в ячейку новое логическое значение</param>
        public static void SetBoolValue(string tableName, string parameterName, int number,
            string chosenParameter, bool value)
        {
            ITable table = _rastr.Tables.Item(tableName);
            ICol columnItem = table.Cols.Item(chosenParameter);

            int index = GetIndexByNumber(tableName, parameterName, number);
            columnItem.set_ZN(index, value);
        }

        /// <summary>
        /// Включить/отключить выбранную линию
        /// </summary>
        /// <param name="ipNumber">Номер начала ветви</param>
        /// <param name="iqNumber">Номер конца ветви</param>
        /// <param name="npNumber">Номер параллельности ветви</param>
        /// <param name="state">Состояние ветви</param>
        /// <exception cref="Exception"></exception>
        public static void ChangeVetvState(int ipNumber, int iqNumber, int npNumber, bool state)
        {
            ITable table = _rastr.Tables.Item("vetv");

            ICol ipColumnItem = table.Cols.Item("ip");
            ICol iqColumnItem = table.Cols.Item("iq");
            ICol npColumnItem = table.Cols.Item("np");
            ICol staColumnItem = table.Cols.Item("sta");

            for (int index = 0; index < table.Count; index++)
            {
                if ((ipColumnItem.get_ZN(index) == ipNumber) &&
                    (iqColumnItem.get_ZN(index) == iqNumber))
                {
                    if (npNumber == 0)
                    {
                        staColumnItem.set_ZN(index, state);
                        break;
                    }
                    else if (npColumnItem.get_ZN(index) == npNumber)
                    {
                        staColumnItem.set_ZN(index, state);
                        break;
                    }   
                }
            }
            //throw new Exception($"Ветвь {ipNumber} - {iqNumber} ({npNumber}) не найдена.");
        }

        // TODO: Неуниверсальная функция, которая работает только с определенной траекторией. Поменять на выбор генераторов вручную или модифицировать
        /// <summary>
        /// Возвращает список генераторов исследуемой или влияющей станции
        /// </summary>
        /// <param name="isPlantResearched">True - получить список генераторов исследуемой станции, false - влияющей станции</param>
        /// <returns>Список генераторов станции</returns>
        public static List<int> GetResRegGenLists(bool isPlantResearched)
        {
            ITable table = _rastr.Tables.Item("ut_node");

            ICol nyColumnItem = table.Cols.Item("ny");
            ICol pgColumnItem = table.Cols.Item("pg");

            List<int> genList = new List<int>();

            for (int index = 0; index < table.Count; index++)
            {
                if ((pgColumnItem.get_ZN(index) > 0) && (isPlantResearched == true))
                {
                    genList.Add(nyColumnItem.get_ZN(index));
                }
                else if ((pgColumnItem.get_ZN(index) < 0) && (isPlantResearched == false))
                {
                    genList.Add(nyColumnItem.get_ZN(index));
                }
            }

            return genList;
        }

        /// <summary>
        /// Получить числовое значение из любой ячейки любой таблицы
        /// </summary>
        /// <param name="tableName">Название таблицы, в которой ведется поиск</param>
        /// <param name="parameterName">Название параметра, по которому ищется индекс</param>
        /// <param name="number">Номер узла (может быть так же номером сечения)</param>
        /// <param name="chosenParameter">Любой параметр, значение из ячеек которого нужно получить
        /// (например, модуль текущего напряжения в узле)</param>
        /// <returns>Числовое значение из ячейки таблицы</returns>
        public static double GetValue(string tableName, string parameterName,
            int number, string chosenParameter)
        {
            ITable table = _rastr.Tables.Item(tableName);
            ICol columnItem = table.Cols.Item(chosenParameter);

            int index = GetIndexByNumber(tableName, parameterName, number);
            return columnItem.get_ZN(index);
        }

        /// <summary>
        /// Получить строковое значение из любой ячейки любой таблицы
        /// </summary>
        /// <param name="tableName">Название таблицы, в которой ведется поиск</param>
        /// <param name="parameterName">Название параметра, по которому ищется индекс</param>
        /// <param name="number">Номер узла (может быть так же номером сечения)</param>
        /// <param name="chosenParameter">Любой параметр, значение из ячеек которого нужно получить
        /// (например, модуль текущего напряжения в узле)</param>
        /// <returns>Строковое значение из ячейки таблицы</returns>
        public static string GetStringValue(string tableName, string parameterName,
            int number, string chosenParameter)
        {
            ITable table = _rastr.Tables.Item(tableName);
            ICol columnItem = table.Cols.Item(chosenParameter);

            int index = GetIndexByNumber(tableName, parameterName, number);
            return columnItem.get_ZN(index);
        }

        /// <summary>
        /// Проверяет режим
        /// </summary>
        /// <returns> Возвращает true - расчет завершен успешно, false - аварийное завершение расчета</returns>
        public static bool IsRegimeOK()
        {
            var statusRgm = _rastr.rgm("");

            if (statusRgm == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Рассчитывает режим
        /// </summary>
        public static void Regime()
        {
            _rastr.rgm("");
        }

        /// <summary>
        /// Записывает желаемое значение в любую ячейку таблицы
        /// </summary>
        /// <param name="tableName">Название таблицы, в которой ведется поиск</param>
        /// <param name="parameterName">Название параметра, по которому ищется индекс</param>
        /// <param name="number">Номер узла (может быть так же номером сечения)</param>
        /// <param name="chosenParameter">Любой параметр, значение из ячеек которого нужно получить
        /// (например, модуль текущего напряжения в узле)</param>
        /// <param name="value">Записываемое в ячейку новое значение</param>
        public static void SetValue(string tableName, string parameterName, int number,
            string chosenParameter, double value)
        {
            ITable table = _rastr.Tables.Item(tableName);
            ICol columnItem = table.Cols.Item(chosenParameter);

            int index = GetIndexByNumber(tableName, parameterName, number);
            columnItem.set_ZN(index, value);
        }

        /// <summary>
        /// Заполняет список номерами узлов/сечений из загруженного файла
        /// </summary>
        /// <param name="tableName">Название таблицы, в которой ведется поиск</param>
        /// <param name="parameterName">Название параметра, по которому ищется индекс></param>
        /// <returns>Список значений рассматриваемого параметра</returns>
        /// <exception cref="Exception"></exception>
        public static List<int> FillNumbersListFromRastr(string tableName, string parameterName)
        {
            List<int> ListOfNumbersFromRastr = new List<int>();

            try
            {
                ITable table = _rastr.Tables.Item(tableName);
                ICol column = table.Cols.Item(parameterName);

                for (int index = 0; index < table.Count; index++)
                {
                    ListOfNumbersFromRastr.Add(column.get_ZN(index));
                }

                return ListOfNumbersFromRastr;
            }
            catch (Exception exeption)
            {
                throw new Exception(exeption.Message);
            }
        }

        /// <summary>
        /// Заполняет список именами контролируемых параметров
        /// </summary>
        /// <returns>Список имен контролируемых параметров</returns>
        /// <exception cref="Exception"></exception>
        public static List<string> FillListOfControlPararmeters()
        {
            List<string> ListOfControlPararmetr = new List<string>();

            try
            {
                ITable table = _rastr.Tables.Item("ots_val");
                ICol column = table.Cols.Item("name");

                for (int index = 0; index < table.Count; index++)
                {
                    ListOfControlPararmetr.Add(column.get_ZN(index));
                }

                return ListOfControlPararmetr;
            }
            catch (Exception exeption)
            {
                throw new Exception(exeption.Message);
            }
        }

        /// <summary>
        /// Шаг назад по траектории
        /// </summary>
        public static void StepBack()
        {
            ITable table = _rastr.Tables.Item("ut_common");
            ICol columnItem = table.Cols.Item("kfc"); //шаг утяжеления

            int index = table.Count - 1; //индекс - самая последняя строка
            double step = columnItem.get_ZN(index);

            columnItem.set_Z(index, -step);

            RastrRetCode kd;

            // шаг утяжеления
            kd = _rastr.step_ut("z");
            if (((kd == 0) && (_rastr.ut_Param[ParamUt.UT_ADD_P] == 0))
                || _rastr.ut_Param[ParamUt.UT_TIP] == 1)
            {
                _rastr.AddControl(-1, ""); //Добавить строку в таблицу значений контролируемых величин
            }

            columnItem.set_Z(index, step);
        }

        /// <summary>
        /// Первичная проверка для всех факторов-сечений:
        ///     Если реакции нет, тогда фактору присваивается Реакция = 0
        ///     Если реакция есть, она рассчитывается для этого сечения-ВФ
        /// </summary>
        /// <param name="factorList">Список влияющих факторов для проведения проверки</param>
        /// <param name="researchingPlantGenerators">Список генераторов исследуемой электростанции</param>
        /// <param name="rstFileName">Имя rst-файла для загрузки</param>
        public static void PrimaryCheckForSectionReaction(List<InfluentFactorBase> factorList,
            List<int> researchingPlantGenerators, string rstFileName)
        {
            string shablonRst = Settings.shablonPaths[ShablonFileType.Scheme];

            ITable tableForIncrement = _rastr.Tables.Item("ut_node");
            ICol columnForStatement = tableForIncrement.Cols.Item("sta");

            // Приращения для всех сечений-факторов выставляем равным нулю (иначе говоря, их строки отключаются)
            foreach (var factor in factorList)
            {
                if (factor is SectionFactor) // Если ВФ это сечение
                {
                    for (int j = 0; j < ((SectionFactor)factor).RegulatingGeneratorsList.Count; j++)
                    {
                        SetBoolValue("ut_node", "ny", ((SectionFactor)factor).RegulatingGeneratorsList[j], "sta", true);
                    }
                }
            }

            foreach (var factor in factorList)
            {
                if (factor is SectionFactor) // Если ВФ это сечение
                {
                    double initialSectionPower = Math.Round(GetValue("sechen", "ns", factor.NumberFromRastr, "psech"), 0);

                    // Расчёт шага утяжеления только основными генераторами (после первого шага цикл прерывается)
                    if (_rastr.ut_Param[ParamUt.UT_FORM_P] == 0)
                    {
                        _rastr.Tables.Item("ut_common").Cols.Item("tip").Z[0] = 0;
                        _rastr.ut_FormControl();
                        _rastr.ClearControl();
                        RastrRetCode kod = _rastr.step_ut("i");
                        if (kod == 0)
                        {
                            RastrRetCode kd;
                            do
                            {
                                // Шаг утяжеления
                                kd = _rastr.step_ut("z");
                                if (((kd == 0) && (_rastr.ut_Param[ParamUt.UT_ADD_P] == 0)) 
                                    || _rastr.ut_Param[ParamUt.UT_TIP] == 1)
                                {
                                    _rastr.AddControl(-1, "");
                                }
                                // Шаг утяжеления
                                double sectionPowerAfterStep = Math.Round(
                                    GetValue("sechen", "ns", factor.NumberFromRastr, "psech"), 0); // Значение перетока после шага утяжеления

                                StepBack(); // Шагаем назад (возвращаемся к исходному режиму)

                                if (sectionPowerAfterStep == initialSectionPower) // Если переток НЕ изменился после шага утяжеления, то приращения остаются равными нулю
                                {
                                    ((SectionFactor)factor).Reaction = 0;
                                    return;
                                }
                                else // Если переток меняется, значит есть зависимость -> нужно рассчитать приращение
                                {
                                    ((SectionFactor)factor).Reaction = sectionPowerAfterStep - initialSectionPower; // Расчёт реакции 

                                    // Обратно включаем номера влияющих генераторов в таблице приращений
                                    for (int i = 0; i < ((SectionFactor)factor).RegulatingGeneratorsList.Count; i++)
                                    {
                                        SetBoolValue("ut_node", "ny", ((SectionFactor)factor).RegulatingGeneratorsList[i], "sta", false);
                                    }

                                    // Выключаем номера генераторов исследуемой станции
                                    for (int i = 0; i < researchingPlantGenerators.Count; i++)
                                    {
                                        SetBoolValue("ut_node", "ny", researchingPlantGenerators[i], "sta", true);
                                    }

                                    if (((SectionFactor)factor).Reaction > 0) // Если реакция положительная, установить приращение влияющим генераторам со знаком "-"
                                    {
                                        for (int i = 0; i < ((SectionFactor)factor).RegulatingGeneratorsList.Count; i++)
                                        {
                                            SetValue("ut_node", "ny", ((SectionFactor)factor).RegulatingGeneratorsList[i], "pg", -1);
                                        }
                                    }
                                    else if (((SectionFactor)factor).Reaction < 0) // Если реакция отрицательная, установить приращение влияющим генераторам со знаком "+"
                                    {
                                        for (int i = 0; i < ((SectionFactor)factor).RegulatingGeneratorsList.Count; i++)
                                        {
                                            SetValue("ut_node", "ny", ((SectionFactor)factor).RegulatingGeneratorsList[i], "pg", 1);
                                        }
                                    }
                                }
                                break; // Прерываем после одного шага
                            }
                            while (kd == 0);
                        }
                    }

                    // Траектория готова. Делаем по ней шаг, чтобы рассчитать коэффициент влияния влияющих генераторов на сечение-ВФ

                    //double powerFlowBeforeStep;
                    double powerFlowAfterStep;

                    // Загружаем исходный режим
                    LoadFile(rstFileName, shablonRst);

                    if (_rastr.ut_Param[ParamUt.UT_FORM_P] == 0)
                    {
                        //_rastr.Tables.Item("ut_common").Cols.Item("tip").Z[0] = 0;
                        _rastr.ut_FormControl();
                        _rastr.ClearControl();
                        RastrRetCode kod = _rastr.step_ut("i");
                        if (kod == 0)
                        {
                            RastrRetCode kd;
                            do
                            {
                                kd = _rastr.step_ut("z");
                                if (((kd == 0) && (_rastr.ut_Param[ParamUt.UT_ADD_P] == 0)) || _rastr.ut_Param[ParamUt.UT_TIP] == 1)
                                {
                                    _rastr.AddControl(-1, "");
                                }
                                // Шаг выполнен. Определяем реакцию

                                powerFlowAfterStep = GetValue("sechen", "ns", factor.NumberFromRastr, "psech"); // Фиксируем переток ПОСЛЕ 
                                StepBack();

                                double InfluentCoeff = ((powerFlowAfterStep - initialSectionPower) / ((SectionFactor)factor).RegulatingGeneratorsList.Count);
                                double compensationPower = ((SectionFactor)factor).Reaction / InfluentCoeff;
                                double increment = compensationPower / ((SectionFactor)factor).RegulatingGeneratorsList.Count;

                                // Присваиваем приращение в таблице траектории

                                for (int i = 0; i < ((SectionFactor)factor).RegulatingGeneratorsList.Count; i++) // Для каждого генератора данного сечения
                                {
                                    SetValue("ut_node", "ny", ((SectionFactor)factor).RegulatingGeneratorsList[i], "pg", increment); // Приращение генерации пересчитанное
                                }
                                break;
                            }
                            while (kd == 0);
                        }
                    }
                }
            }

            for (int i = 0; i < tableForIncrement.Count; i++)
            {
                columnForStatement.set_Z(i, false);
            }

            // Опять исходный режим
            var a = RastrSupplier.IsRegimeOK();

            LoadFile(rstFileName, shablonRst);
        }

        /// <summary>
        /// Функция, которая определяет значение перетока в момент нарушения динамической усточивости
        /// </summary>
        /// <param name="researchingPlantGenerators">Список генераторов исследуемой электростанции</param>
        /// <param name="powerInSection">Текущий переток в сечении</param>
        /// <param name="stepCounter">Номер шага утяжеления</param>
        /// <returns>Возвращает текст для протокола, который выводится на графический интерфейс</returns>
        public static double PowerInSectionCounter(List<int> researchingPlantGenerators,
            double powerInSection, int stepCounter)
        {
            double incrementSum = 0; // Сумма приращений генераторов

            for (int i = 0; i < researchingPlantGenerators.Count; i++)
            {
                incrementSum += GetValue("ut_node", "ny", researchingPlantGenerators[i], "pg");
            }
            double powerInSectionWhenStabilityIsOK = powerInSection - incrementSum;

            MainLayout.console.Log($"Предельный по динамической устойчивости переток составил {powerInSectionWhenStabilityIsOK} МВт.");

            return powerInSectionWhenStabilityIsOK;
        }


        /// <summary>
        /// Функция для определения нарушения динамической устойчивости при минимальном составе ВГО
        /// </summary>
        /// <param name="powerInSection">Текущий переток в сечении</param>
        /// <param name="stepCounter">Номер шага утяжеления</param>
        /// <returns>Возвращает текст для протокола, который выводится на графический интерфейс</returns>
        public static void IfStabilityLossWithMinGen(double powerInSection, int stepCounter)
        {
            MainLayout.console.Log($"Динамическая устойчивость нарушается даже при минимальной загрузке генераторов.");
        }
    }
}
