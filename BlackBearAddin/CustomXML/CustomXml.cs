using NetOffice.ExcelApi;
using NetOffice.OfficeApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace ExcelPythonAddIn.CustomXML
{
    class CustomXml
    {

        public static void Set(Workbook wb, String key, String value)
        {
            //Delete old item if it exists
            var partsEnum = wb.CustomXMLParts.GetEnumerator();
            while (partsEnum.MoveNext())
            {
                var xmlString = partsEnum.Current.XML.ToString();
                var element = new XmlDocument();
                element.LoadXml(xmlString);
                if(element.DocumentElement.Name == key) partsEnum.Current.Delete();
            }
            var keyElement = new XElement(key, value);
            wb.CustomXMLParts.Add(keyElement.ToString());
        }

        public static string Get(Workbook wb, String key)
        {
            var partsEnum = wb.CustomXMLParts.GetEnumerator();
            while (partsEnum.MoveNext())
            {
                var xmlString = partsEnum.Current.XML.ToString();
                var element = new XmlDocument();
                element.LoadXml(xmlString);
                if (element.DocumentElement.Name == key) return element.DocumentElement.InnerText;
            }
            return "";
        }

        public static Dictionary<string, string> GetAll(Workbook wb, string prefix)
        {
            var output = new Dictionary<string, string>();
            var partsEnum = wb.CustomXMLParts.GetEnumerator();
            while (partsEnum.MoveNext())
            {
                var xmlString = partsEnum.Current.XML.ToString();
                var element = new XmlDocument();
                element.LoadXml(xmlString);
                var elementName = element.DocumentElement.Name;
                if (elementName.Substring(0, Math.Min(prefix.Length, elementName.Length)) == prefix)
                {
                    output.Add(element.DocumentElement.Name.Substring(prefix.Length), element.DocumentElement.InnerText);
                }
            }
            return output;
        }
        
    }


}
