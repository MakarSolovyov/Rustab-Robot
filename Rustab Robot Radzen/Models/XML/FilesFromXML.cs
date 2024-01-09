using Rustab_Robot_Radzen.Models.InfluentFactors;
using System.Xml.Serialization;

namespace Rustab_Robot_Radzen.Models.XML
{
    public class FilesFromXML : XMLManager
    {
        public string CalcType { get; set; }

        public string CalcTask { get; set; }

        public string RstFilePath { get; set; }

        public string SchemeExcelFilePath { get; set; }

        public List<string> ScnFilePathList { get; set; }

        public string DfwFilePath { get; set; }

        public string SchFilePath { get; set; }

        public string KprFilePath { get; set; }

        public string Ut2FilePath { get; set; }

        public string ScnInfoFilePath { get; set; }

        public string PPredInfoFilePath { get; set; }

        public List<InfluentFactorBase> InfluentFactorList { get; set; }

        public override string FileXMLPath { get; set; } = Path.Combine(Environment.CurrentDirectory, "RustabBotFileNames.xml");

        public FilesFromXML()
        { }

        public FilesFromXML(string calcType, string calcTask, string rstFilePath,
            string schemeExcelFilePath, List<string> scnFilePathList,
            string dfwFilePath, string schFilePath, string kprFilePath,
            string ut2FilePath, List<InfluentFactorBase> influentFactorList, string scnInfoFilePath, string pPredInfoFilePath)
        {
            CalcType = calcType;

            CalcTask = calcTask;

            RstFilePath = rstFilePath;

            SchemeExcelFilePath = schemeExcelFilePath;

            ScnFilePathList = scnFilePathList;

            DfwFilePath = dfwFilePath;

            SchFilePath = schFilePath;

            KprFilePath = kprFilePath;

            Ut2FilePath = ut2FilePath;

            ScnInfoFilePath = scnInfoFilePath;

            PPredInfoFilePath = pPredInfoFilePath;

            InfluentFactorList = influentFactorList;
        }

        public override FilesFromXML UploadXMLFile(string fileXMLPath)
        {
            XmlSerializer formatter = new XmlSerializer(typeof(FilesFromXML));

            using (FileStream fs = new FileStream(fileXMLPath, FileMode.OpenOrCreate))
            {
                FilesFromXML? xmlFileNames = formatter.Deserialize(fs) as FilesFromXML;

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

            XmlSerializer formatter = new XmlSerializer(typeof(FilesFromXML));

            using (FileStream fs = new FileStream(fileXMLPath, FileMode.OpenOrCreate))
            {
                formatter.Serialize(fs, xmlFileNames);
            }
        }
    }
}
