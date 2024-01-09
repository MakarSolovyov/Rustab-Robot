using Rustab_Robot_Radzen.Models.InfluentFactors;
using System.Xml.Serialization;
using System;
using System.Runtime.Serialization;
using Rustab_Robot_Radzen.Shared;

namespace Rustab_Robot_Radzen.Models.XML
{
    public abstract class XMLManager
    {
        public abstract string FileXMLPath { get; set; }

        public abstract XMLManager UploadXMLFile(string fileXMLPath);

        public abstract void SaveXMLFile(XMLManager xmlFileNames, string fileXMLPath);

        public static bool IsFileXMLExists(string fileXMLPath)
        {
            if (File.Exists(fileXMLPath))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
