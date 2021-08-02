using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using MySqlConnector;
using CsvHelper;

namespace QueryMysqlEveryFiveMinute
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                DateTime currentTime = DateTime.Now;
                string PathOfDesktop = @"C:\Users\MyUser\Desktop";
                if (!Directory.Exists($"{PathOfDesktop}\\POWER_DATA"))
                {
                    Directory.CreateDirectory($"{PathOfDesktop}\\POWER_DATA");
                }
                if (currentTime.Minute % 5 == 0)
                {
                    _logger.LogInformation($"Query DB to CSV File...");
                    try
                    {
                        using (var connection = new MySqlConnection("Server=127.0.0.1;User ID=root;Password=root;Database=icontrol_chenya"))
                        {
                            List<PowerCSV> results = new List<PowerCSV>();
                            connection.Open();

                            using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=159 and eventTime between '{currentTime.AddHours(-8).AddMinutes(-5).ToString("yyyy-MM-dd HH:mm:ss")}' and '{currentTime.AddHours(-8).ToString("yyyy-MM-dd HH:mm:ss")}'", connection))
                            using (var reader = command.ExecuteReader())
                                while (reader.Read())
                                    results.Add(new PowerCSV { bayname = "LINE_1510", value = reader.GetDouble(0), timestamp = reader.GetDateTime(1).AddHours(8) });
                            using (var writer = new StreamWriter($"{PathOfDesktop}\\POWER_DATA\\L1510-{currentTime.ToString("yyMMdd_HHmm")}.csv"))
                            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                            {
                                csv.WriteRecords(results);
                            }

                        }
                        using (var connection = new MySqlConnection("Server=127.0.0.1;User ID=root;Password=root;Database=icontrol_chenya"))
                        {
                            List<PowerCSV> results = new List<PowerCSV>();
                            connection.Open();

                            using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=1064 and eventTime between '{currentTime.AddHours(-8).AddMinutes(-5).ToString("yyyy-MM-dd HH:mm:ss")}' and '{currentTime.AddHours(-8).ToString("yyyy-MM-dd HH:mm:ss")}'", connection))
                            using (var reader = command.ExecuteReader())
                                while (reader.Read())
                                    results.Add(new PowerCSV { bayname = "MP3", value = reader.GetDouble(0), timestamp = reader.GetDateTime(1).AddHours(8) });
                            using (var writer = new StreamWriter($"{PathOfDesktop}\\POWER_DATA\\MP3-{currentTime.ToString("yyMMdd_HHmm")}.csv"))
                            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                            {
                                csv.WriteRecords(results);
                            }

                        }
                        using (var connection = new MySqlConnection("Server=127.0.0.1;User ID=root;Password=root;Database=icontrol_chenya"))
                        {
                            List<PowerCSV> results = new List<PowerCSV>();
                            connection.Open();

                            using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=1068 and eventTime between '{currentTime.AddHours(-8).AddMinutes(-5).ToString("yyyy-MM-dd HH:mm:ss")}' and '{currentTime.AddHours(-8).ToString("yyyy-MM-dd HH:mm:ss")}'", connection))
                            using (var reader = command.ExecuteReader())
                                while (reader.Read())
                                    results.Add(new PowerCSV { bayname = "MP4", value = reader.GetDouble(0), timestamp = reader.GetDateTime(1).AddHours(8) });
                            using (var writer = new StreamWriter($"{PathOfDesktop}\\POWER_DATA\\MP4-{currentTime.ToString("yyMMdd_HHmm")}.csv"))
                            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                            {
                                csv.WriteRecords(results);
                            }

                        }
                    }
                    catch(Exception e)
                    {
                        _logger.LogError(e.Message);
                    }
                }
                //_logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                await Task.Delay(1000, stoppingToken);
            }
        }
    }
}
