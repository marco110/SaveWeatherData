using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace SaveWeatherData
{
    class Program
    {

        static void Main(string[] args)
        {
            Stopwatch watch = new Stopwatch();
            watch.Start();
            DirectoryInfo folder = new DirectoryInfo(@"D:\我的文档\Visual Studio 2012\Projects\GetWeatherData\GetWeatherData\bin\Debug\results");
            List<string> errorLog = new List<string>();
            var files = folder.GetFiles("*.json");

            int filecount = 1;
            foreach (var file in files)
            {               
                string fileName = file.FullName;

                List<string> groupsData = File.ReadLines(fileName, Encoding.UTF8).ToList(); //new List<string>() { "{city:'北京',tqInfo:[{ymd:'2011-12-01',bWendu:'℃',yWendu:'11℃',tianqi:'小雨~阴',fengxiang:'东北风',fengli:'微风'},{ymd:'2011-12-02',bWendu:'16℃',yWendu:'9℃',tianqi:'多云',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-03',bWendu:'13℃',yWendu:'7℃',tianqi:'多云~阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-04',bWendu:'12℃',yWendu:'8℃',tianqi:'小雨',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-05',bWendu:'12℃',yWendu:'9℃',tianqi:'小雨',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-06',bWendu:'12℃',yWendu:'10℃',tianqi:'小雨',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-07',bWendu:'11℃',yWendu:'9℃',tianqi:'小雨',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-08',bWendu:'9℃',yWendu:'7℃',tianqi:'小雨',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-09',bWendu:'9℃',yWendu:'7℃',tianqi:'小雨',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-10',bWendu:'8℃',yWendu:'5℃',tianqi:'小雨~阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-11',bWendu:'10℃',yWendu:'7℃',tianqi:'阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-12',bWendu:'11℃',yWendu:'5℃',tianqi:'多云~阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-13',bWendu:'11℃',yWendu:'5℃',tianqi:'多云~阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-14',bWendu:'9℃',yWendu:'6℃',tianqi:'阴~小雨',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-15',bWendu:'11℃',yWendu:'8℃',tianqi:'阴~小雨',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-16',bWendu:'10℃',yWendu:'7℃',tianqi:'阴~小雨',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-17',bWendu:'10℃',yWendu:'7℃',tianqi:'阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-18',bWendu:'9℃',yWendu:'3℃',tianqi:'阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-19',bWendu:'10℃',yWendu:'6℃',tianqi:'阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-20',bWendu:'9℃',yWendu:'6℃',tianqi:'阴~小雨',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-21',bWendu:'9℃',yWendu:'7℃',tianqi:'小雨',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-22',bWendu:'9℃',yWendu:'7℃',tianqi:'小雨',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-23',bWendu:'11℃',yWendu:'6℃',tianqi:'阴~多云',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-24',bWendu:'13℃',yWendu:'5℃',tianqi:'晴~多云',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-25',bWendu:'10℃',yWendu:'5℃',tianqi:'多云~阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-26',bWendu:'10℃',yWendu:'6℃',tianqi:'多云~阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-27',bWendu:'9℃',yWendu:'7℃',tianqi:'阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-28',bWendu:'10℃',yWendu:'6℃',tianqi:'阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-29',bWendu:'10℃',yWendu:'7℃',tianqi:'阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-30',bWendu:'9℃',yWendu:'7℃',tianqi:'阴',fengxiang:'无持续风向',fengli:'微风'},{ymd:'2011-12-31',bWendu:'11℃',yWendu:'5℃',tianqi:'多云',fengxiang:'无持续风向',fengli:'微风'},{}],maxWendu:'16（2011-12-02）',minWendu:'3（2011-12-18）',avgbWendu:'11',avgyWendu:'7'}" };//File.ReadAllLines(inputPath, UTF8Encoding.UTF8).ToList();
                List<List<WeatherInfo>> weatherInfos = new List<List<WeatherInfo>>();

                int count = 0;

                foreach (var data in groupsData)
                {
                    var jObject = (JObject)JsonConvert.DeserializeObject(data);

                    try
                    {
                        List<WeatherInfo> weatherInfo = ConverJObjectToWeatherInfo(jObject);
                        count += weatherInfo.Count;
                        weatherInfos.Add(weatherInfo);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        errorLog.Add(jObject["city"].ToString()+" " + jObject["tqInfo"]["ymd"].ToString());
                    }
                    Console.WriteLine("转换第 "+filecount+" 个文件  " + count);
                }

                groupsData.Clear();
                WriteXls(@"D:\我的文档\Visual Studio 2012\Projects\GetWeatherData\GetWeatherData\bin\Debug\results\" + file.Name + ".xlsx", weatherInfos, count,filecount);
                filecount++;
                File.WriteAllLines(@"D:\我的文档\Visual Studio 2012\Projects\GetWeatherData\GetWeatherData\bin\Debug\results\" + file.Name + "errorlog.txt", errorLog);
                Console.WriteLine("完事！");
            }           

            watch.Stop();
            Console.WriteLine("总耗时： " + watch.Elapsed);
            Console.WriteLine(DateTime.Now);
            Console.ReadKey();
        }


        private static List<WeatherInfo> ConverJObjectToWeatherInfo(JObject jObject)
        {
            //var jObject = (JObject)JsonConvert.DeserializeObject(result);
            List<WeatherInfo> weatherList = new List<WeatherInfo>();

            int weatherCount = jObject["tqInfo"].Count();

            for (int i = 0; i < weatherCount; i++)
            {
                var result = jObject["tqInfo"][i];

                if (result.First != null)
                {
                    try
                    {
                        string bWendu = result["bWendu"].ToString();
                        string yWendu = result["yWendu"].ToString();
                        int tempbWndu = 0, tempyWndu = 0;
                        if (Regex.IsMatch(bWendu, @"\d+"))
                        {
                            tempbWndu = Convert.ToInt32(Regex.Replace(bWendu, @"(\d+).", @"${1}"));
                        }
                        else
                        {
                            tempbWndu = Convert.ToInt32(Regex.Replace(yWendu, @"(\d+).", @"${1}"));
                        }
                        if (Regex.IsMatch(yWendu, @"\d+"))
                        {
                            tempyWndu = Convert.ToInt32(Regex.Replace(yWendu, @"(\d+).", @"${1}"));
                        }
                        else
                        {
                            tempyWndu = tempbWndu;
                        }


                        int temperature = (tempbWndu + tempyWndu) / 2;

                        string weather = result["tianqi"].ToString();
                        if (Regex.IsMatch(weather, "晴")) weather = "晴";
                        if (Regex.IsMatch(weather, "多云") || Regex.IsMatch(weather, "阴")) weather = "多云";
                        if (Regex.IsMatch(weather, "雾")) weather = "雾";
                        if (Regex.IsMatch(weather, "雨")) weather = "雨";
                        if (Regex.IsMatch(weather, "雪")) weather = "雪";                        

                        string wind = string.Empty;
                        string tempWind = result["fengxiang"].ToString();
                        if (Regex.IsMatch(tempWind, "东")) wind = "东";
                        if (Regex.IsMatch(tempWind, "西")) wind = "西";
                        if (Regex.IsMatch(tempWind, "南")) wind = "南";
                        if (Regex.IsMatch(tempWind, "北")) wind = "北";
                        if (Regex.IsMatch(tempWind, "东南")) wind = "东南";
                        if (Regex.IsMatch(tempWind, "东北")) wind = "东北";
                        if (Regex.IsMatch(tempWind, "西南")) wind = "西南";
                        if (Regex.IsMatch(tempWind, "西北")) wind = "西北";
                        if (Regex.IsMatch(tempWind, "无") || Regex.IsMatch(tempWind, "微")) wind = "无";

                        WeatherInfo weatherInfo = new WeatherInfo(jObject["city"].ToString(), result["ymd"].ToString(), temperature, weather, wind);
                        weatherList.Add(weatherInfo);
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        Console.WriteLine(jObject["city"].ToString()+ result["ymd"].ToString());
                    }
                }
            }
            return weatherList;
        }

        public static bool WriteXls(string filename, List<List<WeatherInfo>> weatherInfos, int count,int filecount)
        {

            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            _Workbook book = xls.Workbooks.Add(Missing.Value); //创建一张表，一张表可以包含多个sheet

            //如果表已经存在，可以用下面的命令打开
            //_Workbook book = xls.Workbooks.Open(filename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            _Worksheet sheet;
            xls.Visible = false;//设置Excel后台运行
            xls.DisplayAlerts = false;//设置不显示确认修改提示

            //for (int i = 1; i < 2; i++)//循环创建并写入数据到sheet
            int row = 1;
            {
                try
                {
                    sheet = (_Worksheet)book.Worksheets.get_Item(1);
                }
                catch (Exception ex)//不存在就增加一个sheet
                {
                    Console.WriteLine(ex.Message);
                    sheet = (_Worksheet)book.Worksheets.Add(Missing.Value, book.Worksheets[book.Sheets.Count], 1, Missing.Value);
                    Console.WriteLine("已添加sheet");
                }
                //sheet.Name = "第" + 1 + "页";
                for (int i = 0; i < weatherInfos.Count; i++)
                {
                    for (int j = 0; j < weatherInfos[i].Count; j++)
                    {
                        sheet.Cells[row, 1] = weatherInfos[i][j].CountyName;
                        sheet.Cells[row, 2] = weatherInfos[i][j].Date;
                        sheet.Cells[row, 3] = weatherInfos[i][j].Temperature;
                        sheet.Cells[row, 4] = weatherInfos[i][j].Weather;
                        sheet.Cells[row, 5] = weatherInfos[i][j].Wind;

                        Console.WriteLine("写入第 " + filecount + " 个文件  " + row + " / " + count);
                        row++;
                    }
                    weatherInfos.RemoveAt(i);
                    i--;
                }
            }

            book.SaveAs(filename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //book.Save();

            book.Close(false, Missing.Value, Missing.Value);
            xls.Quit();
            sheet = null;
            book = null;
            xls = null;
            GC.Collect();
            return true;
        }

        private static Array ReadXls(string filename, int index)
        {

            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();

            _Workbook book = xls.Workbooks.Open(filename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            _Worksheet sheet;
            xls.Visible = false;
            xls.DisplayAlerts = false;

            try
            {
                sheet = (_Worksheet)book.Worksheets.get_Item(index);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }

            Console.WriteLine(sheet.Name);
            int row = sheet.UsedRange.Rows.Count;
            int col = sheet.UsedRange.Columns.Count;
            Array value = (Array)sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[row, col]).Cells.Value2;

            book.Save();
            book.Close(false, Missing.Value, Missing.Value);
            xls.Quit();
            sheet = null;
            book = null;
            xls = null;
            GC.Collect();
            return value;
        }
    }
}
