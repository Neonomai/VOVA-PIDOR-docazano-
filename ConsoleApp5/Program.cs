using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        // Убедитесь, что файл 'result.json' существует в директории с программой
        string json = File.ReadAllText("result.json");
        var jsonObject = JObject.Parse(json);

        // Проверьте, что 'messages' действительно является массивом в вашем JSON
        var messages = jsonObject["messages"] as JArray;

        var data = new List<LogEntry>();

        foreach (var message in messages)
        {
            var textArray = message["text"] as JArray;
            string hostname = null;
            string date = null;
            string level = null;
            string problem = null;

            foreach (var item in textArray)
            {
                if (item.Type == JTokenType.Object)
                {
                    var itemObj = (JObject)item;
                    if (itemObj["type"] != null && itemObj["type"].ToString() == "code")
                    {
                        // Предполагаем, что 'itemObj["text"]' содержит вашу JSON-строку
                        string jsonText = itemObj["text"].ToString().Trim();
                        Console.WriteLine("JSON перед разбором: " + jsonText); // Для отладки
                        if (jsonText.Contains("caller"))
                        {
                            var codeText = JObject.Parse(jsonText);
                            problem = codeText["msg"].ToString();

                            //if (jsonText.Contains("")) { }
                        }
                        else if (!(jsonText[0] == 'z' && jsonText[1] == 'b'))
                        {
                            problem = jsonText;
                        }
                        else
                        {
                            hostname = jsonText;
                        }
                    }
                    else if (itemObj["type"] != null && itemObj["type"].ToString() == "bold")
                    {
                        // Предполагаем, что 'itemObj["text"]' содержит вашу JSON-строку
                        string jsonText = itemObj["text"].ToString().Trim();
                        Console.WriteLine("JSON перед разбором: " + jsonText); // Для отладки
                        if (jsonText.Contains("Warning") || jsonText.Contains("Average") || jsonText.Contains("Error"))
                        {
                            level = jsonText.ToString();

                            //if (jsonText.Contains("")) { }
                        }
                        else if (jsonText.Contains("🗓"))
                        {
                            date += "🕒 ";
                            date += jsonText.ToString();
                        }
                        
                    }
                }
              
            }

            data.Add(new LogEntry
            {
                Hostname = hostname,
                DateTime = date,
                Problem = problem,
                Level = level
            });
        }

        SaveToExcel(data);
    }

    public class LogEntry
    {
        public string Hostname { get; set; }
        public string DateTime { get; set; }
        public string Problem { get; set; }
        public string Level { get; set; }
    }

    private static void SaveToExcel(List<LogEntry> data)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Logs");

            // Заголовки столбцов
            worksheet.Cells[1, 1].Value = "Hostname";
            worksheet.Cells[1, 2].Value = "DateTime";
            worksheet.Cells[1, 3].Value = "Problem";
            worksheet.Cells[1, 4].Value = "Level";

            // Заполнение данных
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = data[i].Hostname;
                worksheet.Cells[i + 2, 2].Value = data[i].DateTime;
                worksheet.Cells[i + 2, 3].Value = data[i].Problem;
                worksheet.Cells[i + 2, 4].Value = data[i].Level;
            }

            // Сохранение файла
            var file = new FileInfo("logs.xlsx");
            package.SaveAs(file);
        }
    }
}
