using CsvHelper.Configuration.Attributes;

namespace Rustab_Robot_Radzen.Models.DSI
{
    public class CSVCalcResult
    {
        [Name("t")]
        public double Time { get; set; }
        [Name("U_BoGES")]
        public double UBoGES { get; set; }
        [Name("Delta_BoGES_G1")]
        public double DeltaBoGESGen1 { get; set; }
        [Name("Delta_BoGES_G3")]
        public double DeltaBoGESGen3 { get; set; }
        [Name("Pg_BoGES_G1")]
        public double PgBoGESGen1 { get; set; }
        [Name("Pg_BoGES_G3")]
        public double PgBoGESGen3 { get; set; }
        [Name("Delta_Haranor_G3")]
        public double DeltaHaranorGen3 { get; set; }
        [Name("Delta_Belovo_G1")]
        public double DeltaBelovoGen1 { get; set; }
        [Name("Delta_Kras_G5")]
        public double DeltaKrasGen5 { get; set; }
        [Name("Delta_SSHGES_G9")]
        public double DeltaSSHGESGen9 { get; set; }
        [Name("Delta_Bratsk_G4")]
        public double DeltaBratskGen4 { get; set; }
        [Name("Delta_Irkutsk_G2")]
        public double DeltaIrkutskGen2 { get; set; }
        [Name("Delta_Gusi_G5")]
        public double DeltaGusiGen5 { get; set; }
        [Name("Delta_Chita_G4")]
        public double DeltaChitaGen4 { get; set; }
    }
}
