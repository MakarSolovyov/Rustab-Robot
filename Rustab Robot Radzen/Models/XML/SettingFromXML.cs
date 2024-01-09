using System.Xml.Serialization;

namespace Rustab_Robot_Radzen.Models.XML
{
    public class SettingFromXML : XMLManager
    {
        public string CalcResultFolder { get; set; }

        public string RankResultFolder { get; set; }

        public double CalcDSITimevalue { get; set; }

        public double CalcTimevalue { get; set; }

        public override string FileXMLPath { get; set; } = Path.Combine(Environment.CurrentDirectory, "RustabBotSettings.xml");

        public SettingFromXML()
        { }

        public SettingFromXML(string calcResultFolder, string rankResultFolder, double calcDSITimevalue, double calcTimevalue)
        {
            CalcResultFolder = calcResultFolder;
            RankResultFolder = rankResultFolder;
            CalcDSITimevalue = calcDSITimevalue;
            CalcTimevalue = calcTimevalue;
        }

        public override SettingFromXML UploadXMLFile(string fileXMLPath)
        {
            XmlSerializer formatter = new XmlSerializer(typeof(SettingFromXML));

            using (FileStream fs = new FileStream(fileXMLPath, FileMode.OpenOrCreate))
            {
                SettingFromXML? xmlFileNames = formatter.Deserialize(fs) as SettingFromXML;

                if (xmlFileNames != null)
                {
                    return xmlFileNames;
                }
                else
                {
                    throw new Exception($"Ошибка загрузки файла \"{fileXMLPath}\".");
                }
            }
        }

        public override void SaveXMLFile(XMLManager xmlFileNames, string fileXMLPath)
        {
            if (IsFileXMLExists(fileXMLPath))
            {
                File.Delete(fileXMLPath);
            }

            XmlSerializer formatter = new XmlSerializer(typeof(SettingFromXML));

            using (FileStream fs = new FileStream(fileXMLPath, FileMode.OpenOrCreate))
            {
                formatter.Serialize(fs, xmlFileNames);
            }
        }
    }
}
