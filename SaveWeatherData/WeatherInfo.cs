namespace SaveWeatherData
{
    public class WeatherInfo
    {
        public string CountyName { get; set; }
        public string Date { get; set; }
        public int Temperature { get; set; }
        public string Weather { get; set; }
        public string Wind { get; set; }

        public WeatherInfo() { }

        public WeatherInfo(string countyName, string date, int temperature, string weather, string wind)
        {
            this.CountyName = countyName;
            this.Date = date;
            this.Temperature = temperature;
            this.Weather = weather;
            this.Wind = wind;
        }
    }
}
