using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Xml;

namespace MrResXInserter
{
    class Program
    {
        static void Main(string[] args)
        {
            // if (args.Length < 2) {
            //     throw new Exception("There must at least be specified an input path and an output path as parametres!");
            // }
            string inputPath = @"C:\temp\resx-files-637921770930399585.xlsx"; // args[0];
            string outputPath = @"C:\temp\output"; // args[1];

            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                pck.Load(File.OpenRead(inputPath));
                
                for (int i = 0; i < pck.Workbook.Worksheets.Count(); i++) {
                    var ws = pck.Workbook.Worksheets[i];

                    if (ws.Dimension == null) {
                        return;
                    }

                    var file = Directory.GetFiles(outputPath, "*.*", SearchOption.AllDirectories).Where(x => x.ToLower().EndsWith(ws.Name.Replace("-", ".") + ".resx")).FirstOrDefault();
                    
                    if (file == null) {
                        continue;
                    }

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(file);
                    XmlNodeList translations = xmlDoc.GetElementsByTagName("data");
                    
                    var start = ws.Dimension.Start;
                    var end = ws.Dimension.End;

                    /* Rows */
                    for (int row = start.Row; row <= end.Row; row++)
                    {
                        string source = "";

                        /* Cells */
                        for (int col = start.Column; col <= end.Column; col++)
                        {
                            string text = ws.Cells[row, col].Text;

                            if (source.Length == 0) {
                                source = text;
                            } else {
                                InsertTranslation(translations, source, text);
                            }
                        }
                    }

                    // var ib = xmlDoc.SelectNodes("//data/*");

                    // xmlDoc.GetElementsByTagName("data")
                    //     .Cast<XmlNode>()
                    //     .OrderBy(x => x.Attributes["name"].Value)
                    //     .ToList();

                    xmlDoc.Save(file);
                }
            }


                // var ws = pck.Workbook.Worksheets[sheet];




            

            // List<string> files = Directory.GetFiles(inputPath, "*.*", SearchOption.AllDirectories).ToList().FindAll(x => x.ToLower().EndsWith(".resx"));

            // if (language != null) {
            //     files = files.FindAll(x => x.ToLower().Replace(".resx", "").EndsWith("." + language));
            // }
            
            // string output = outputPath + @"\resx-files-" + (language != null ? (language + "-") : "") + DateTime.Now.Ticks + ".xlsx";

            // using (var pck = new OfficeOpenXml.ExcelPackage())
            // {
            //     foreach (var file in files) {
            //         XmlDocument xmlDoc = new XmlDocument();
            //         xmlDoc.Load(file);

            //         int row = 1;
                    
            //         string fileName = file.Split(@"\")[file.Split(@"\").Count() - 1];
            //         string sheet = fileName.ToLower().Replace(".", "-").Replace("-resx", "");

            //         /* Avoid error due to max length of a sheet title */
            //         if (sheet.Length > 31) {
            //             sheet = sheet.Split("-")[0].Substring(0, 28) + "-" + sheet.Split("-")[1];
            //         }

            //         pck.Workbook.Worksheets.Add(sheet);
            //         var ws = pck.Workbook.Worksheets[sheet];  
                       
            //         XmlNodeList translations = xmlDoc.GetElementsByTagName("data");
                    
            //         foreach (XmlNode translation in translations) {
            //             var source = translation.Attributes["name"].Value;
            //             var trans = translation.SelectSingleNode("value").InnerText;
            //             ws.Cells[row, 1].Value = source.Trim();
            //             ws.Cells[row, 2].Value = trans.Trim();
            //             row++;
            //         }

            //         ws.Cells.AutoFitColumns();
            //     }

            //     pck.SaveAs(new FileInfo(output));
            // }            
        }

        private static void FileCheck() {

        }

        private static void InsertTranslation(XmlNodeList translations, string key, string newTrans) {
            foreach (XmlNode node in translations) {
                if (node.Attributes["name"].Value == key) {
                    node.SelectSingleNode("value").InnerText = newTrans;
                    break;
                }
            }
        }
    }
}

                    