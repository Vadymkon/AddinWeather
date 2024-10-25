using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using RestSharp;
using Newtonsoft.Json.Linq;
using System.Xml.Linq;
using System.Threading.Tasks;

namespace TestAddinWeather
{
    public partial class Weather
    {
        private readonly string apiKey = "dcccfefbd932d021573b9a94645ac4f5";
        private readonly string apiKeyCrypto = "CF4415CB-E39E-4B31-94D1-732434BC56CB";
        private readonly string baseUrlWeather = "http://api.openweathermap.org/data/2.5/forecast";
        private readonly string baseUrlCrypto = "https://rest.coinapi.io/v1/trades/BITSTAMP_SPOT_BTC_USD/history";

        public async Task<JObject> GetCryptoDataAsync()
        {
            var client = new RestClient();
            var request = new RestRequest(baseUrlCrypto, Method.Get);
            request.AddParameter("date", "2020-07-10");
            request.AddHeader("Accept", "text/plain");
            request.AddHeader("X-CoinAPI-Key", apiKey);
            RestResponse response = await client.ExecuteAsync(request);

            string result = response.Content;

            return JObject.Parse(result);
        }

        public void InsertCryptoDataToExcel(Excel.Worksheet worksheet, JObject cryptoData)
        {
            worksheet.Cells[1, 1] = "Currency";
            worksheet.Cells[1, 2] = "Rate";

            var currency = cryptoData["asset_id_base"].ToString();
            var rate = cryptoData["rate"].ToObject<double>();

            worksheet.Cells[2, 1] = currency;
            worksheet.Cells[2, 2] = rate;
        }

        public async Task<JArray> GetWeatherDataAsync(double lat, double lon)
        {
            var client = new RestClient(baseUrlWeather);
            var request = new RestRequest();

            request.AddParameter("lat", lat);
            request.AddParameter("lon", lon);
            request.AddParameter("appid", apiKey);
            request.AddParameter("units", "metric");

            var response = await client.ExecuteAsync(request);
            string result = response.Content;

            var json = JObject.Parse(result);
            return (JArray)json["list"];
        }

        public void InsertWeatherDataToExcel(Excel.Worksheet worksheet, JArray weatherData)
        {
            int startRow = 2;
            worksheet.Cells[1, 1] = "Date";
            worksheet.Cells[1, 2] = "Temperature (°C)";
            worksheet.Cells[1, 3] = "Feels Like (°C)";
            worksheet.Cells[1, 4] = "Weather Description";

            foreach (var dataPoint in weatherData)
            {
                var date = dataPoint["dt_txt"].ToString();
                var temp = dataPoint["main"]["temp"].ToObject<double>();
                var feelsLike = dataPoint["main"]["feels_like"].ToObject<double>();
                var description = dataPoint["weather"][0]["description"].ToString();

                worksheet.Cells[startRow, 1] = date;
                worksheet.Cells[startRow, 2] = temp;
                worksheet.Cells[startRow, 3] = feelsLike;
                worksheet.Cells[startRow, 4] = description;
                startRow++;
            }

            if (startRow > 2)
            {
                if (worksheet.AutoFilter != null)
                {
                    worksheet.AutoFilterMode = false;
                }

                Excel.Range usedRange = worksheet.UsedRange;
                usedRange.AutoFilter(1);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No data to filter.");
            }
        }

        private void Weather_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private async void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorkSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            if (activeWorkSheet != null)
            {
                double latitude = 50.4501;
                double longitude = 30.5234;

                var weatherData = await GetWeatherDataAsync(latitude, longitude);

                InsertWeatherDataToExcel(activeWorkSheet, weatherData);
            }
        }

        private async void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorkSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            if (activeWorkSheet != null)
            {
                var cryptoData = await GetCryptoDataAsync();

                InsertCryptoDataToExcel(activeWorkSheet, cryptoData);
            }
        }
    }
}
