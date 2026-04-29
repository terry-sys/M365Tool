using System;
using System.Globalization;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace Office365CleanupTool.Services
{
    public sealed class WeatherSnapshot
    {
        public bool IsSuccess { get; set; }

        public string Location { get; set; } = string.Empty;

        public int WeatherCode { get; set; }

        public double TemperatureCelsius { get; set; }

        public DateTime RetrievedAtLocal { get; set; }

        public string ErrorMessage { get; set; } = string.Empty;
    }

    public sealed class WeatherService
    {
        private static readonly HttpClient HttpClient = new()
        {
            Timeout = TimeSpan.FromSeconds(12)
        };

        public async Task<WeatherSnapshot> GetCurrentWeatherByIpAsync()
        {
            try
            {
                string ipJson = await HttpClient.GetStringAsync("https://ipapi.co/json/");
                using JsonDocument ipDocument = JsonDocument.Parse(ipJson);
                JsonElement ipRoot = ipDocument.RootElement;

                if (!TryGetDouble(ipRoot, "latitude", out double latitude) ||
                    !TryGetDouble(ipRoot, "longitude", out double longitude))
                {
                    return new WeatherSnapshot
                    {
                        IsSuccess = false,
                        ErrorMessage = "无法获取当前 IP 所在位置。"
                    };
                }

                string city = GetString(ipRoot, "city");
                string region = GetString(ipRoot, "region");
                string country = GetString(ipRoot, "country_name");
                string location = $"{city}{(string.IsNullOrWhiteSpace(region) ? string.Empty : ", " + region)}{(string.IsNullOrWhiteSpace(country) ? string.Empty : ", " + country)}".Trim(' ', ',');

                string weatherUrl =
                    $"https://api.open-meteo.com/v1/forecast?latitude={latitude.ToString(CultureInfo.InvariantCulture)}&longitude={longitude.ToString(CultureInfo.InvariantCulture)}&current=temperature_2m,weather_code&timezone=auto";
                string weatherJson = await HttpClient.GetStringAsync(weatherUrl);

                using JsonDocument weatherDocument = JsonDocument.Parse(weatherJson);
                JsonElement current = weatherDocument.RootElement.GetProperty("current");

                double temperature = current.GetProperty("temperature_2m").GetDouble();
                int weatherCode = current.GetProperty("weather_code").GetInt32();

                return new WeatherSnapshot
                {
                    IsSuccess = true,
                    Location = string.IsNullOrWhiteSpace(location) ? "Unknown" : location,
                    TemperatureCelsius = temperature,
                    WeatherCode = weatherCode,
                    RetrievedAtLocal = DateTime.Now
                };
            }
            catch (Exception ex)
            {
                return new WeatherSnapshot
                {
                    IsSuccess = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        public static string ToWeatherText(int weatherCode, UiLanguage language)
        {
            if (language == UiLanguage.English)
            {
                return weatherCode switch
                {
                    0 => "Clear",
                    1 or 2 => "Partly cloudy",
                    3 => "Overcast",
                    45 or 48 => "Fog",
                    51 or 53 or 55 => "Drizzle",
                    61 or 63 or 65 => "Rain",
                    71 or 73 or 75 => "Snow",
                    80 or 81 or 82 => "Rain showers",
                    95 or 96 or 99 => "Thunderstorm",
                    _ => "Unknown"
                };
            }

            return weatherCode switch
            {
                0 => "晴",
                1 or 2 => "多云",
                3 => "阴",
                45 or 48 => "雾",
                51 or 53 or 55 => "毛毛雨",
                61 or 63 or 65 => "雨",
                71 or 73 or 75 => "雪",
                80 or 81 or 82 => "阵雨",
                95 or 96 or 99 => "雷暴",
                _ => "未知"
            };
        }

        private static bool TryGetDouble(JsonElement root, string property, out double value)
        {
            value = 0;
            if (!root.TryGetProperty(property, out JsonElement element))
            {
                return false;
            }

            if (element.ValueKind == JsonValueKind.Number)
            {
                value = element.GetDouble();
                return true;
            }

            if (element.ValueKind == JsonValueKind.String)
            {
                return double.TryParse(
                    element.GetString(),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out value);
            }

            return false;
        }

        private static string GetString(JsonElement root, string property)
        {
            if (!root.TryGetProperty(property, out JsonElement element))
            {
                return string.Empty;
            }

            return element.ValueKind == JsonValueKind.String
                ? element.GetString() ?? string.Empty
                : element.ToString();
        }
    }
}
