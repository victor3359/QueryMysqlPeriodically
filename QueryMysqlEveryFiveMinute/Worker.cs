using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

using MySqlConnector;
using OfficeOpenXml;

namespace QueryMysqlEveryFiveMinute
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        private bool oldState_Daily;
        private bool raiseFlag_Daily = false;

        private bool oldState_Monthly;
        private bool raiseFlag_Monthly = false;

        private string DesktopPath;
        public Worker(ILogger<Worker> logger, string DesktopPath)
        {
            _logger = logger;
            this.DesktopPath = DesktopPath;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (!Directory.Exists($"{DesktopPath}\\POWER_DATA"))
            {
                Directory.CreateDirectory($"{DesktopPath}\\POWER_DATA");
            }
            if (!Directory.Exists($"{DesktopPath}\\POWER_DATA\\Daily"))
            {
                Directory.CreateDirectory($"{DesktopPath}\\POWER_DATA\\Daily");
            }
            if (!Directory.Exists($"{DesktopPath}\\POWER_DATA\\Monthly"))
            {
                Directory.CreateDirectory($"{DesktopPath}\\POWER_DATA\\Monthly");
            }
        }

        private void AddDailySheets(MySqlConnection connection, ExcelPackage ep, string bayName, string sheetName, DateTime currentTime, int dbIndex)
        {
            int index = 2;
            ep.Workbook.Worksheets.Add($"{bayName}_{sheetName}");
            ExcelWorksheet st = ep.Workbook.Worksheets[$"{bayName}_{sheetName}"];
            st.Cells[1, 1].LoadFromText($"BAYNAME,VALUE,TIMESTAMP");
            st.Column(3).Style.Numberformat.Format = @"yyyy/MM/dd HH:mm:ss.000";
            using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint={dbIndex} and eventTime between '{currentTime.AddHours(-8).AddHours(-24).ToString("yyyy-MM-dd HH:mm:ss")}' and '{currentTime.AddHours(-8).ToString("yyyy-MM-dd HH:mm:ss")}'", connection))
            using (var reader = command.ExecuteReader())
                while (reader.Read())
                    st.Cells[index++, 1].LoadFromText($"{bayName},{reader.GetDouble(0)},{reader.GetDateTime(1).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")}");
            st.Cells.AutoFitColumns();
        }
        private void AddMonthlySheets(MySqlConnection connection, ExcelPackage ep, string bayName, string sheetName, DateTime currentTime, int dbIndex)
        {
            int index = 2;
            ep.Workbook.Worksheets.Add($"{bayName}_{sheetName}");
            ExcelWorksheet st = ep.Workbook.Worksheets[$"{bayName}_{sheetName}"];
            st.Cells[1, 1].LoadFromText($"BAYNAME,VALUE,TIMESTAMP");
            st.Column(3).Style.Numberformat.Format = @"yyyy/MM/dd HH:mm:ss.000";
            using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint={dbIndex} and eventTime between '{currentTime.AddHours(-8).AddMonths(-1).ToString("yyyy-MM-dd HH:mm:ss")}' and '{currentTime.AddHours(-8).ToString("yyyy-MM-dd HH:mm:ss")}'", connection))
            using (var reader = command.ExecuteReader())
                while (reader.Read())
                    st.Cells[index++, 1].LoadFromText($"{bayName},{reader.GetDouble(0)},{reader.GetDateTime(1).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")}");
            st.Cells.AutoFitColumns();
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                DateTime currentTime = DateTime.Now;

                oldState_Daily = raiseFlag_Daily;
                raiseFlag_Daily = currentTime.Hour == 2 ? true : false;

                oldState_Monthly = raiseFlag_Monthly;
                raiseFlag_Monthly = currentTime.Day == 1 ? true : false;

                //Daily Report
                if (oldState_Daily == false && raiseFlag_Daily == true)
                {
                    _logger.LogInformation($"Query DB to Excel File...");
                    try
                    {
                        using (ExcelPackage ep = new ExcelPackage())
                        {
                            using (var connection = new MySqlConnection("Server=127.0.0.1;User ID=root;Password=root;Database=icontrol_chenya"))
                            {
                                connection.Open();

                                AddDailySheets(connection, ep, "LINE1510", "FWD", currentTime, 161);
                                AddDailySheets(connection, ep, "LINE1510", "REV", currentTime, 159);
                                AddDailySheets(connection, ep, "DTR1650", "FWD", currentTime, 165);
                                AddDailySheets(connection, ep, "DTR1650", "REV", currentTime, 163);
                                AddDailySheets(connection, ep, "DTR1660", "FWD", currentTime, 169);
                                AddDailySheets(connection, ep, "DTR1660", "REV", currentTime, 167);
                                AddDailySheets(connection, ep, "MP1", "FWD", currentTime, 1058);
                                AddDailySheets(connection, ep, "MP1", "REV", currentTime, 1056);
                                AddDailySheets(connection, ep, "MP2", "FWD", currentTime, 1062);
                                AddDailySheets(connection, ep, "MP2", "REV", currentTime, 1060);
                                AddDailySheets(connection, ep, "MP3", "FWD", currentTime, 1066);
                                AddDailySheets(connection, ep, "MP3", "REV", currentTime, 1064);
                                AddDailySheets(connection, ep, "MP4", "FWD", currentTime, 1070);
                                AddDailySheets(connection, ep, "MP4", "REV", currentTime, 1068);
                                AddDailySheets(connection, ep, "TIE", "FWD", currentTime, 1074);
                                AddDailySheets(connection, ep, "TIE", "REV", currentTime, 1072);
                                AddDailySheets(connection, ep, "FEEDER_11", "FWD", currentTime, 994);
                                AddDailySheets(connection, ep, "FEEDER_11", "REV", currentTime, 992);
                                AddDailySheets(connection, ep, "FEEDER_12", "FWD", currentTime, 998);
                                AddDailySheets(connection, ep, "FEEDER_12", "REV", currentTime, 996);
                                AddDailySheets(connection, ep, "FEEDER_13", "FWD", currentTime, 1002);
                                AddDailySheets(connection, ep, "FEEDER_13", "REV", currentTime, 1000);
                                AddDailySheets(connection, ep, "FEEDER_14", "FWD", currentTime, 1006);
                                AddDailySheets(connection, ep, "FEEDER_14", "REV", currentTime, 1004);
                                AddDailySheets(connection, ep, "FEEDER_15", "FWD", currentTime, 1010);
                                AddDailySheets(connection, ep, "FEEDER_15", "REV", currentTime, 1008);
                                AddDailySheets(connection, ep, "FEEDER_16", "FWD", currentTime, 1014);
                                AddDailySheets(connection, ep, "FEEDER_16", "REV", currentTime, 1012);
                                AddDailySheets(connection, ep, "FEEDER_17", "FWD", currentTime, 1018);
                                AddDailySheets(connection, ep, "FEEDER_17", "REV", currentTime, 1016);
                                AddDailySheets(connection, ep, "FEEDER_18", "FWD", currentTime, 1022);
                                AddDailySheets(connection, ep, "FEEDER_18", "REV", currentTime, 1020);
                                AddDailySheets(connection, ep, "FEEDER_21", "FWD", currentTime, 1026);
                                AddDailySheets(connection, ep, "FEEDER_21", "REV", currentTime, 1024);
                                AddDailySheets(connection, ep, "FEEDER_22", "FWD", currentTime, 1030);
                                AddDailySheets(connection, ep, "FEEDER_22", "REV", currentTime, 1028);
                                AddDailySheets(connection, ep, "FEEDER_23", "FWD", currentTime, 1034);
                                AddDailySheets(connection, ep, "FEEDER_23", "REV", currentTime, 1032);
                                AddDailySheets(connection, ep, "FEEDER_24", "FWD", currentTime, 1038);
                                AddDailySheets(connection, ep, "FEEDER_24", "REV", currentTime, 1036);
                                AddDailySheets(connection, ep, "FEEDER_25", "FWD", currentTime, 1042);
                                AddDailySheets(connection, ep, "FEEDER_25", "REV", currentTime, 1040);
                                AddDailySheets(connection, ep, "FEEDER_26", "FWD", currentTime, 1046);
                                AddDailySheets(connection, ep, "FEEDER_26", "REV", currentTime, 1044);
                                AddDailySheets(connection, ep, "FEEDER_27", "FWD", currentTime, 1050);
                                AddDailySheets(connection, ep, "FEEDER_27", "REV", currentTime, 1048);
                                AddDailySheets(connection, ep, "FEEDER_28", "FWD", currentTime, 1054);
                                AddDailySheets(connection, ep, "FEEDER_28", "REV", currentTime, 1052);

                                //EVENTLIST
                                int index = 2;
                                ep.Workbook.Worksheets.Add("EVENTLIST");
                                ExcelWorksheet EVENTLIST = ep.Workbook.Worksheets["EVENTLIST"];
                                EVENTLIST.Cells[1, 1].LoadFromText($"EVENT,STATE,TIMESTAMP");
                                EVENTLIST.Column(3).Style.Numberformat.Format = @"yyyy/MM/dd HH:mm:ss.000";
                                using (var command = new MySqlCommand($"select T3, state, eventTime from digitalevents inner join points on digitalevents.points_idPoint = points.idPoint where eventTime between '{currentTime.AddHours(-8).AddHours(-24).ToString("yyyy-MM-dd HH:mm:ss")}' and '{currentTime.AddHours(-8).ToString("yyyy-MM-dd HH:mm:ss")}'", connection))
                                using (var reader = command.ExecuteReader())
                                    while (reader.Read())
                                        if(reader.GetString(0) != "") EVENTLIST.Cells[index++, 1].LoadFromText($"{reader.GetString(0)},{reader.GetString(1)},{reader.GetDateTime(2).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")}");
                                EVENTLIST.Cells.AutoFitColumns();

                                //Save ExcelFile
                                FileInfo fi = new FileInfo($"{DesktopPath}\\POWER_DATA\\Daily\\CHENYA-{currentTime.ToString("yyyyMMdd")}_DailyReport.xlsx");
                                ep.SaveAs(fi);
                            }
                        }
                    }
                    catch(Exception e)
                    {
                        _logger.LogError(e.Message);
                    }
                }
                //Monthly Report
                if (oldState_Monthly == false && raiseFlag_Monthly == true)
                {
                    using (ExcelPackage ep = new ExcelPackage())
                    {
                        using (var connection = new MySqlConnection("Server=127.0.0.1;User ID=root;Password=root;Database=icontrol_chenya"))
                        {
                            connection.Open();

                            AddMonthlySheets(connection, ep, "LINE1510", "FWD", currentTime, 161);
                            AddMonthlySheets(connection, ep, "LINE1510", "REV", currentTime, 159);
                            AddMonthlySheets(connection, ep, "DTR1650", "FWD", currentTime, 165);
                            AddMonthlySheets(connection, ep, "DTR1650", "REV", currentTime, 163);
                            AddMonthlySheets(connection, ep, "DTR1660", "FWD", currentTime, 169);
                            AddMonthlySheets(connection, ep, "DTR1660", "REV", currentTime, 167);
                            AddMonthlySheets(connection, ep, "MP1", "FWD", currentTime, 1058);
                            AddMonthlySheets(connection, ep, "MP1", "REV", currentTime, 1056);
                            AddMonthlySheets(connection, ep, "MP2", "FWD", currentTime, 1062);
                            AddMonthlySheets(connection, ep, "MP2", "REV", currentTime, 1060);
                            AddMonthlySheets(connection, ep, "MP3", "FWD", currentTime, 1066);
                            AddMonthlySheets(connection, ep, "MP3", "REV", currentTime, 1064);
                            AddMonthlySheets(connection, ep, "MP4", "FWD", currentTime, 1070);
                            AddMonthlySheets(connection, ep, "MP4", "REV", currentTime, 1068);
                            AddMonthlySheets(connection, ep, "TIE", "FWD", currentTime, 1074);
                            AddMonthlySheets(connection, ep, "TIE", "REV", currentTime, 1072);
                            AddMonthlySheets(connection, ep, "FEEDER_11", "FWD", currentTime, 994);
                            AddMonthlySheets(connection, ep, "FEEDER_11", "REV", currentTime, 992);
                            AddMonthlySheets(connection, ep, "FEEDER_12", "FWD", currentTime, 998);
                            AddMonthlySheets(connection, ep, "FEEDER_12", "REV", currentTime, 996);
                            AddMonthlySheets(connection, ep, "FEEDER_13", "FWD", currentTime, 1002);
                            AddMonthlySheets(connection, ep, "FEEDER_13", "REV", currentTime, 1000);
                            AddMonthlySheets(connection, ep, "FEEDER_14", "FWD", currentTime, 1006);
                            AddMonthlySheets(connection, ep, "FEEDER_14", "REV", currentTime, 1004);
                            AddMonthlySheets(connection, ep, "FEEDER_15", "FWD", currentTime, 1010);
                            AddMonthlySheets(connection, ep, "FEEDER_15", "REV", currentTime, 1008);
                            AddMonthlySheets(connection, ep, "FEEDER_16", "FWD", currentTime, 1014);
                            AddMonthlySheets(connection, ep, "FEEDER_16", "REV", currentTime, 1012);
                            AddMonthlySheets(connection, ep, "FEEDER_17", "FWD", currentTime, 1018);
                            AddMonthlySheets(connection, ep, "FEEDER_17", "REV", currentTime, 1016);
                            AddMonthlySheets(connection, ep, "FEEDER_18", "FWD", currentTime, 1022);
                            AddMonthlySheets(connection, ep, "FEEDER_18", "REV", currentTime, 1020);
                            AddMonthlySheets(connection, ep, "FEEDER_21", "FWD", currentTime, 1026);
                            AddMonthlySheets(connection, ep, "FEEDER_21", "REV", currentTime, 1024);
                            AddMonthlySheets(connection, ep, "FEEDER_22", "FWD", currentTime, 1030);
                            AddMonthlySheets(connection, ep, "FEEDER_22", "REV", currentTime, 1028);
                            AddMonthlySheets(connection, ep, "FEEDER_23", "FWD", currentTime, 1034);
                            AddMonthlySheets(connection, ep, "FEEDER_23", "REV", currentTime, 1032);
                            AddMonthlySheets(connection, ep, "FEEDER_24", "FWD", currentTime, 1038);
                            AddMonthlySheets(connection, ep, "FEEDER_24", "REV", currentTime, 1036);
                            AddMonthlySheets(connection, ep, "FEEDER_25", "FWD", currentTime, 1042);
                            AddMonthlySheets(connection, ep, "FEEDER_25", "REV", currentTime, 1040);
                            AddMonthlySheets(connection, ep, "FEEDER_26", "FWD", currentTime, 1046);
                            AddMonthlySheets(connection, ep, "FEEDER_26", "REV", currentTime, 1044);
                            AddMonthlySheets(connection, ep, "FEEDER_27", "FWD", currentTime, 1050);
                            AddMonthlySheets(connection, ep, "FEEDER_27", "REV", currentTime, 1048);
                            AddMonthlySheets(connection, ep, "FEEDER_28", "FWD", currentTime, 1054);
                            AddMonthlySheets(connection, ep, "FEEDER_28", "REV", currentTime, 1052);

                            FileInfo fi = new FileInfo($"{DesktopPath}\\POWER_DATA\\Monthly\\CHENYA-{currentTime.ToString("yyyyMM")}_MonthlyReport.xlsx");
                            ep.SaveAs(fi);
                        }
                    }
                }
                await Task.Delay(1000, stoppingToken);
            }
        }
    }
}
