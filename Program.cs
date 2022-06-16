using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Xml;

namespace MrResXtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2) {
                throw new Exception("There must at least be specified an input path and an output path as parametres!");
            }
            string inputPath = args[0];
            string outputPath = args[1];
            string language = null;

            if (args.Length > 2) {
                language = args[2];
            }
            

            List<string> files = Directory.GetFiles(inputPath, "*.*", SearchOption.AllDirectories).ToList().FindAll(x => x.ToLower().EndsWith(".resx"));

            if (language != null) {
                files = files.FindAll(x => x.ToLower().Replace(".resx", "").EndsWith("." + language));
            }
            
            string output = outputPath + @"\resx-files-" + (language != null ? (language + "-") : "") + DateTime.Now.Ticks + ".xlsx";

            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                foreach (var file in files) {
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(file);

                    int row = 1;
                    
                    string fileName = file.Split(@"\")[file.Split(@"\").Count() - 1];
                    string sheet = fileName.ToLower().Replace(".", "-").Replace("-resx", "");

                    /* Avoid error due to max length of a sheet title */
                    if (sheet.Length > 31) {
                        sheet = sheet.Split("-")[0].Substring(0, 28) + "-" + sheet.Split("-")[1];
                    }

                    pck.Workbook.Worksheets.Add(sheet);
                    var ws = pck.Workbook.Worksheets[sheet];  
                       
                    XmlNodeList translations = xmlDoc.GetElementsByTagName("data");
                    
                    foreach (XmlNode translation in translations) {
                        var source = translation.Attributes["name"].Value;
                        var trans = translation.SelectSingleNode("value").InnerText;
                        ws.Cells[row, 1].Value = source.Trim();
                        ws.Cells[row, 2].Value = trans.Trim();
                        row++;
                    }

                    ws.Cells.AutoFitColumns();
                }

                pck.SaveAs(new FileInfo(output));
            }            
        }
    }
}
