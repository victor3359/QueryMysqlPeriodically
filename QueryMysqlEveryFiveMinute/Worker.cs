using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

using MySqlConnector;
using OfficeOpenXml;
using System.Diagnostics;

namespace ICP_REPORT_SERVICE
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly IOptions<ServiceOptions> _options;

        private bool oldState_Daily;
        private bool raiseFlag_Daily = false;

        private bool oldState_Monthly;
        private bool raiseFlag_Monthly = false;

        private bool oldState_Archive;
        private bool raiseFlag_Archive = false;

        private bool DebugMode = false;
        private string DebugStr;
        public Worker(ILogger<Worker> logger, IOptions<ServiceOptions> options)
        {
            _logger = logger;
            _options = options;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (!Directory.Exists($"{_options.Value.ReportDirectory}\\Report"))
            {
                Directory.CreateDirectory($"{_options.Value.ReportDirectory}\\Report");
            }
            if (!Directory.Exists($"{_options.Value.ReportDirectory}\\Report\\Excel"))
            {
                Directory.CreateDirectory($"{_options.Value.ReportDirectory}\\Report\\Excel");
            }
            if (!Directory.Exists($"{_options.Value.ReportDirectory}\\Report\\Excel\\Daily"))
            {
                Directory.CreateDirectory($"{_options.Value.ReportDirectory}\\Report\\Excel\\Daily");
            }
            if (!Directory.Exists($"{_options.Value.ReportDirectory}\\Report\\Excel\\Monthly"))
            {
                Directory.CreateDirectory($"{_options.Value.ReportDirectory}\\Report\\Excel\\Monthly");
            }
            if (!Directory.Exists(_options.Value.BackupDirectory))
            {
                Directory.CreateDirectory(_options.Value.BackupDirectory);
            }

            CheckDbTableExist_Create();
        }

        private void AddDailySheets(MySqlConnection connection, ExcelPackage ep, string bayName, string sheetName, DateTime currentTime, int dbIndex, bool HV = false)
        {
            int index = 2;
            ep.Workbook.Worksheets.Add($"{bayName}_{sheetName}");
            ExcelWorksheet st = ep.Workbook.Worksheets[$"{bayName}_{sheetName}"];
            if (HV)
            {
                st.Cells[1, 1].LoadFromText($"BAYNAME,VALUE,TIMESTAMP,,MWh,KWh");
            }
            else
            {
                st.Cells[1, 1].LoadFromText($"BAYNAME,VALUE,TIMESTAMP,,kWh");
            }
            st.Column(3).Style.Numberformat.Format = @"yyyy/MM/dd HH:mm:ss.000";
            DateTime Temp = currentTime.AddHours(-32).AddMinutes(-1);
            double td = 0;
            for (int i=0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint={dbIndex} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i+1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (i == 0)
                            {
                                td = reader.GetDouble(0);
                                break;
                            }
                            if (HV)
                            {
                                st.Cells[index++, 1].LoadFromText($"{bayName},{reader.GetDouble(0)},{reader.GetDateTime(1).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")},,{reader.GetDouble(0) - td},{(reader.GetDouble(0) - td) * 1000}");
                            }
                            else
                            {
                                st.Cells[index++, 1].LoadFromText($"{bayName},{reader.GetDouble(0)},{reader.GetDateTime(1).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")},,{reader.GetDouble(0) - td}");
                            }

                            td = reader.GetDouble(0);
                        }
                    st.Cells.AutoFitColumns();
                }
            }
            if (HV)
            {
                st.Cells[1, 8].Value = "當日累積發電";
                st.Cells[1, 9].Value = "MWh";
                st.Cells[2, 9].Value = "kWh";
                st.Cells[1, 10].Formula = $"SUM(E2:E{st.Dimension.End.Row})";
                st.Cells[2, 10].Formula = "J1 * 1000";
            }
            else
            {
                st.Cells[1, 8].Value = "當日累積發電";
                st.Cells[1, 9].Value = "kWh";
                st.Cells[1, 10].Formula = $"SUM(E2:E{st.Dimension.End.Row})";
            }
        }
        private void AddMonthlySheets(MySqlConnection connection, ExcelPackage ep, string bayName, string sheetName, DateTime currentTime, int dbIndex, bool HV = false)
        {
            int index = 2;
            ep.Workbook.Worksheets.Add($"{bayName}_{sheetName}");
            ExcelWorksheet st = ep.Workbook.Worksheets[$"{bayName}_{sheetName}"];
            if (HV)
            {
                st.Cells[1, 1].LoadFromText($"BAYNAME,VALUE,TIMESTAMP,,MWh,kWh");
            }
            else
            {
                st.Cells[1, 1].LoadFromText($"BAYNAME,VALUE,TIMESTAMP,,kWh");
            }
            st.Column(3).Style.Numberformat.Format = @"yyyy/MM/dd HH:mm:ss.000";
            DateTime LastMonth = currentTime.AddMonths(-1);
            DateTime LastMonth_UTC = LastMonth.AddHours(-8);
            DateTime Temp = new DateTime(LastMonth_UTC.Year, LastMonth_UTC.Month, DateTime.DaysInMonth(LastMonth_UTC.Year, LastMonth_UTC.Month) - 1, 16, 0, 0);
            double td = 0;
            for (int i = 0; i < DateTime.DaysInMonth(LastMonth.Year, LastMonth.Month) + 1; i++)
            {
                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint={dbIndex} and eventTime between '{Temp.AddDays(i).ToString("yyyy-MM-dd 16:00")}' and '{Temp.AddDays(i+1).ToString($"yyyy-MM-dd 16:00")}' order by eventTime desc limit 1", connection))
                {
                    _logger.LogInformation($"DATE:\nFROM:{Temp.AddDays(i).ToString("yyyy-MM-dd 16:00")}\nTO:{Temp.AddDays(i + 1).ToString($"yyyy-MM-dd 16:00")}");
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if(i == 0)
                            {
                                td = reader.GetDouble(0);
                                break;
                            }
                            if (HV)
                            {
                                st.Cells[index++, 1].LoadFromText($"{bayName},{reader.GetDouble(0)},{reader.GetDateTime(1).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")},,{reader.GetDouble(0) - td},{(reader.GetDouble(0) - td) * 1000}");
                            }
                            else
                            {
                                st.Cells[index++, 1].LoadFromText($"{bayName},{reader.GetDouble(0)},{reader.GetDateTime(1).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")},,{reader.GetDouble(0) - td}");
                            }
                            
                            td = reader.GetDouble(0);
                        }
                    st.Cells.AutoFitColumns();
                }
            }
            if (HV)
            {
                st.Cells[1, 8].Value = "當日累積發電";
                st.Cells[1, 9].Value = "MWh";
                st.Cells[2, 9].Value = "kWh";
                st.Cells[1, 10].Formula = $"SUM(E2:E{st.Dimension.End.Row})";
                st.Cells[2, 10].Formula = "J1 * 1000";
            }
            else
            {
                st.Cells[1, 8].Value = "當日累積發電";
                st.Cells[1, 9].Value = "kWh";
                st.Cells[1, 10].Formula = $"SUM(E2:E{st.Dimension.End.Row})";
            }

            st.Cells.AutoFitColumns();
        }

        private void CheckDbTableExist_Create()
        {
            string CreateTableQueryString = @"CREATE TABLE IF NOT EXISTS `archive_action_log` (
	                    `IND` INT(11) NOT NULL AUTO_INCREMENT,
	                    `ACTION` VARCHAR(50) NULL DEFAULT NULL COLLATE 'utf8_unicode_ci',
	                    `DESCRIPTION` VARCHAR(100) NULL DEFAULT NULL COLLATE 'utf8_unicode_ci',
	                    `DATETIME` DATETIME NULL DEFAULT NULL,
	                    PRIMARY KEY (`IND`)
                    )
                    COLLATE='utf8_unicode_ci'
                    ENGINE=InnoDB
                    ;";
            using (var connection = new MySqlConnection($"Server={_options.Value.MySQL_IpAddress};User ID={_options.Value.MySQL_User};Password={_options.Value.MySQL_Password};Database={_options.Value.MySQL_DbTable}"))
            {
                connection.Open();
                
                using (var command = new MySqlCommand(CreateTableQueryString, connection))
                {
                    command.CommandTimeout = 6000;
                    command.ExecuteNonQuery();
                }
            }

        }
        private void InsertMsgToDbTable(string Action, string Message)
        {
            DateTime currentTime = DateTime.Now;
            using (var connection = new MySqlConnection($"Server={_options.Value.MySQL_IpAddress};User ID={_options.Value.MySQL_User};Password={_options.Value.MySQL_Password};Database={_options.Value.MySQL_DbTable}"))
            {
                connection.Open();

                using (var command = new MySqlCommand($"Insert into archive_action_log(ACTION,DESCRIPTION,DATETIME) values('{Action}','{Message}','{currentTime.ToString("yyyy-MM-dd HH:mm:ss.000")}')", connection))
                {
                    command.CommandTimeout = 6000;
                    command.ExecuteNonQuery();
                }
            }
        }
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                DateTime currentTime = DateTime.Now;

                oldState_Daily = raiseFlag_Daily;
                raiseFlag_Daily = currentTime.Hour == _options.Value.DailyReportTime ? true : false;

                oldState_Monthly = raiseFlag_Monthly;
                raiseFlag_Monthly = currentTime.Day == 1 ? true : false;

                oldState_Archive = raiseFlag_Archive;
                raiseFlag_Archive = currentTime.Hour == _options.Value.ArchiveTime ? true : false;

                DebugStr = DebugMode ? "_Debug" : "";

                //Daily Report
                if ((oldState_Daily == false && raiseFlag_Daily == true) || DebugMode)
                {
                    _logger.LogInformation($"Query DB to Excel File...");
                    try
                    {
                        using (ExcelPackage ep = new ExcelPackage())
                        {
                            using (var connection = new MySqlConnection($"Server={_options.Value.MySQL_IpAddress};User ID={_options.Value.MySQL_User};Password={_options.Value.MySQL_Password};Database={_options.Value.MySQL_DbTable}"))
                            {
                                connection.Open();

                                AddDailySheets(connection, ep, "LINE1510", "FWD", currentTime, 161, true);
                                AddDailySheets(connection, ep, "LINE1510", "REV", currentTime, 159, true);
                                AddDailySheets(connection, ep, "DTR1650", "FWD", currentTime, 165, true);
                                AddDailySheets(connection, ep, "DTR1650", "REV", currentTime, 163, true);
                                AddDailySheets(connection, ep, "DTR1660", "FWD", currentTime, 169, true);
                                AddDailySheets(connection, ep, "DTR1660", "REV", currentTime, 167, true);
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
                                using (var command = new MySqlCommand($"select T3, state, eventTime from digitalevents inner join points on digitalevents.points_idPoint = points.idPoint where eventTime between '{currentTime.AddHours(-8).AddHours(-26).ToString("yyyy-MM-dd HH:mm:ss")}' and '{currentTime.AddHours(-8).AddHours(-2).ToString("yyyy-MM-dd HH:mm:ss")}'", connection))
                                using (var reader = command.ExecuteReader())
                                    while (reader.Read())
                                        if(reader.GetString(0) != "") EVENTLIST.Cells[index++, 1].LoadFromText($"{reader.GetString(0)},{reader.GetString(1)},{reader.GetDateTime(2).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")}");
                                EVENTLIST.Cells.AutoFitColumns();

                                //Save ExcelFile
                                FileInfo fi = new FileInfo($"{_options.Value.ReportDirectory}\\Report\\Excel\\Daily\\CHENYA-{currentTime.AddDays(-1).ToString("yyyyMMdd")}_DailyReport{DebugStr}.xlsx");
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
                if ((oldState_Monthly == false && raiseFlag_Monthly == true) || DebugMode)
                {
                    using (ExcelPackage ep = new ExcelPackage())
                    {
                        using (var connection = new MySqlConnection($"Server={_options.Value.MySQL_IpAddress};User ID={_options.Value.MySQL_User};Password={_options.Value.MySQL_Password};Database={_options.Value.MySQL_DbTable}"))
                        {
                            connection.Open();

                            AddMonthlySheets(connection, ep, "LINE1510", "FWD", currentTime, 161, true);
                            AddMonthlySheets(connection, ep, "LINE1510", "REV", currentTime, 159, true);
                            AddMonthlySheets(connection, ep, "DTR1650", "FWD", currentTime, 165, true);
                            AddMonthlySheets(connection, ep, "DTR1650", "REV", currentTime, 163, true);
                            AddMonthlySheets(connection, ep, "DTR1660", "FWD", currentTime, 169, true);
                            AddMonthlySheets(connection, ep, "DTR1660", "REV", currentTime, 167, true);
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

                            FileInfo fi = new FileInfo($"{_options.Value.ReportDirectory}\\Report\\Excel\\Monthly\\CHENYA-{currentTime.AddMonths(-1).ToString("yyyyMM")}_MonthlyReport{DebugStr}.xlsx");
                            ep.SaveAs(fi);
                        }
                    }
                }
                //Archive Service
                if((oldState_Archive == false && raiseFlag_Archive == true))
                {
                    _logger.LogInformation("--- Archive Service Start ---");
                    InsertMsgToDbTable("ARCHIVE START", $"DATE: {currentTime.ToString("yyyy-MM-dd")} archive service start.");
                    //Scan Database
                    _logger.LogInformation("Scaning rawdata from MySQL...");
                    using (var connection = new MySqlConnection($"Server={_options.Value.MySQL_IpAddress};User ID={_options.Value.MySQL_User};Password={_options.Value.MySQL_Password};Database={_options.Value.MySQL_DbTable}"))
                    {
                        connection.Open();
                        DateTime QueryDate = currentTime.AddDays(_options.Value.KeepDataAliveDays * -1);
                        bool YesterdayHaveRawdata = true;
                        int AddDay = 0;
                        bool TodayHaveRawdata = false;
                        int ScanIndex_End = 0;
                        int ScanIndex_Start = 0;
                        while (YesterdayHaveRawdata)
                        {
                            _logger.LogInformation($"DATE {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} scaning data.");
                            //insert log to DB
                            InsertMsgToDbTable("DATABASE SCAN START", $"DATE: {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} scaning data.");
                            using (var scan_end = new MySqlCommand($"select idEvent from analogevents where eventTime between '{QueryDate.AddDays(AddDay - 1).ToString("yyyy/MM/dd 16:00:00")}' and '{QueryDate.AddDays(AddDay).ToString("yyyy/MM/dd 16:00:00")}' order by eventTime desc limit 1", connection))
                            {
                                scan_end.CommandTimeout = 6000;
                                using (var result_end = scan_end.ExecuteReader())
                                    while (result_end.Read())
                                    {
                                        TodayHaveRawdata = true;
                                        ScanIndex_End = result_end.GetInt32(0);
                                    }
                            }
                            using (var scan_start = new MySqlCommand($"select idEvent from analogevents where eventTime between '{QueryDate.AddDays(AddDay - 1).ToString("yyyy/MM/dd 16:00:00")}' and '{QueryDate.AddDays(AddDay).ToString("yyyy/MM/dd 16:00:00")}' order by eventTime asc limit 1", connection))
                            {
                                scan_start.CommandTimeout = 6000;
                                using (var result_start = scan_start.ExecuteReader())
                                    while (result_start.Read())
                                    {
                                        ScanIndex_Start = result_start.GetInt32(0);
                                    }
                            }
                            _logger.LogInformation($"DATE {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} scan completed.");
                            //insert log to DB
                            InsertMsgToDbTable("DATABASE SCAN END", $"DATE: {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} scan completed.");
                            if (!TodayHaveRawdata)
                            {
                                InsertMsgToDbTable("NO DATA NEED ARCHIVE", $"DATE: {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} have no data to archive.");
                                YesterdayHaveRawdata = false;
                                _logger.LogInformation("There is no more rawdata to archive, exiting service..."); ;
                                _logger.LogInformation("--- Archive Service Finished ---"); ;
                                InsertMsgToDbTable("ARCHIVE FINISHED", $"DATE: {currentTime.ToString("yyyy-MM-dd")} archive service finished.");
                                break;
                            }
                            _logger.LogInformation("Exporting rawdata from MySQL and saving file...");
                            InsertMsgToDbTable("EXPORT DATA START", $"DATE: {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} exporting data, from ID: {ScanIndex_Start} to {ScanIndex_End}.");
                            try
                            {
                                //Save SQL file
                                FileStream StreamDB = new FileStream($"{_options.Value.BackupDirectory}\\{QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")}_Rawdata.sql", FileMode.Create, FileAccess.Write);
                                using (StreamWriter SW = new StreamWriter(StreamDB))
                                {
                                    ProcessStartInfo proc = new ProcessStartInfo();
                                    string cmd = $" --host={_options.Value.MySQL_IpAddress} --user={_options.Value.MySQL_User} --password={_options.Value.MySQL_Password} {_options.Value.MySQL_DbTable} analogevents --where=\"eventTime between '{QueryDate.AddDays(AddDay - 1).ToString("yyyy/MM/dd 16:00:00")}' and '{QueryDate.AddDays(AddDay).ToString("yyyy/MM/dd 16:00:00")}'\"";
                                    // Configure path for mysqldump.exe
                                    proc.FileName = _options.Value.EXEPATH;
                                    proc.RedirectStandardInput = false;
                                    proc.RedirectStandardOutput = true;
                                    proc.UseShellExecute = false;
                                    proc.WindowStyle = ProcessWindowStyle.Minimized;
                                    proc.Arguments = cmd;
                                    proc.CreateNoWindow = true;
                                    Process p = Process.Start(proc);
                                    SW.Write(p.StandardOutput.ReadToEnd());
                                    p.WaitForExit();
                                    p.Close();
                                    SW.Close();
                                    StreamDB.Close();
                                }
                            }
                            catch (Exception e)
                            {
                                _logger.LogError($"Error occured when exporting data: {e.Message}");
                                InsertMsgToDbTable("ERROR OCCURED WHEN EXPORTING", $"DATE: {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} export failed.");
                                continue;
                            }
                            TodayHaveRawdata = false;
                            _logger.LogInformation($"SaveFile: \"{QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")}_Rawdata.sql\"");
                            _logger.LogInformation($"DATE {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} archive completed.");
                            InsertMsgToDbTable("EXPORT DATA END", $"DATE: {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} export completed and saving file.");

                            //Remove rawdata from database
                            _logger.LogInformation($"Remove \"{QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")}\" rawdata from database.");
                            InsertMsgToDbTable("REMOVE DATA START", $"DATE: {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} removing data.");
                            using (var command = new MySqlCommand($"delete from analogevents where TO_DAYS(eventTime) = TO_DAYS(\"{QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")}\")", connection))
                            {
                                command.CommandTimeout = 6000;
                                command.ExecuteNonQuery();
                            }
                            _logger.LogInformation($"Remove completed.");
                            InsertMsgToDbTable("REMOVE DATA END", $"DATE: {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} remove completed.");
                            AddDay--;
                        }
                    }
                }
                await Task.Delay(1000, stoppingToken);
            }
        }
    }
}
