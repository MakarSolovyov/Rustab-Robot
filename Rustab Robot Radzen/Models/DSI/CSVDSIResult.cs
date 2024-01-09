using CsvHelper.Configuration.Attributes;

namespace Rustab_Robot_Radzen.Models.DSI
{
    public class CSVDSIResult
    {
        [Name("№ п/п")]
        public int ID { get; set; }
        [Name("№ сценария")]
        public int ScnNumber { get; set; }
        [Name("Схема сети")]
        public string Scheme { get; set; }
        [Name("Кол-во ВГО")]
        public int GenNumber { get; set; }
        [Name("Загрузка ГГ, МВт")]
        public double GenPower { get; set; }
        [Name("Вид КЗ")]
        public string KZType { get; set; }
        [Name("Место КЗ")]
        public string KZLocation { get; set; }
        [Name("tликв.КЗ, мс")]
        public double KZElimTime { get; set; }
        [Name("dU, кВ")]
        public double VoltageDrop { get; set; }
        [Name("dP, МВт")]
        public double GenPowerDrop { get; set; }
        [Name("Pпред.пар., МВт")]
        public double PPredPAR { get; set; }
        [Name("delta0, град")]
        public double InitDelta { get; set; }
        [Name("PМДП, МВт")]
        public double MaxPowerFlow { get; set; }
        [Name("Нарушение ДУ")]
        public bool IsStabilityLost { get; set; }
        [Name("КПТ, о.е.")]
        public double ComplexDSI { get; set; }
        [Name("Pпред., МВт")]
        public double? PPred { get; set; }
    }
}
