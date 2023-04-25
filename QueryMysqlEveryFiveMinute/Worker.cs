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
using System.IO.Compression;

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

        private bool oldState_TenDays;
        private bool raiseFlag_TenDays = false;
        private int raiseState_TenDays = 0;

        private bool oldState_Archive;
        private bool raiseFlag_Archive = false;
        private bool Archive_Is_Finished = true;

        // If True, will just archive database.
        private bool ArchiveOnly = false;

        // If True, will just output reports.
        private bool ReportOnly = false;

        // If Debugmode, output file will add debug text.
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
            if (!Directory.Exists($"{_options.Value.ReportDirectory}\\Report\\Excel\\TenDays"))
            {
                Directory.CreateDirectory($"{_options.Value.ReportDirectory}\\Report\\Excel\\TenDays");
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

        private void AddDeviceTempSheets(MySqlConnection connection, ExcelPackage ep, string bayName, string sheetName, DateTime currentTime, int temp1, int temp2)
        {
            int index = 2;
            ep.Workbook.Worksheets.Add($"{bayName}_{sheetName}");
            ExcelWorksheet st = ep.Workbook.Worksheets[$"{bayName}_{sheetName}"];
            st.Cells[1, 1].LoadFromText($"ID,TIMESTAMP,Temp1,Temp2");
            st.Column(2).Style.Numberformat.Format = @"yyyy/MM/dd HH:mm:ss.000";
            DateTime Temp = currentTime.AddHours(-32).AddMinutes(-1);
            for (int i = 0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint={temp1} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0)) break;
                            st.Cells[index++, 1].LoadFromText($"{i + 1},{reader.GetDateTime(1).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")},{reader.GetDouble(0).ToString("C2")}");
                        }
                }
            }
            index = 2;
            for (int i = 0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value from analogevents where points_idPoint={temp2} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0)) break;
                            st.Cells[index++, 4].LoadFromText($"{reader.GetDouble(0).ToString("C2")}");
                        }
                }
            }
            st.Cells.AutoFitColumns();
        }
        private void AddRadiationSheets(MySqlConnection connection, ExcelPackage ep, string bayName, string sheetName, DateTime currentTime, int R1, int R2, int R3, int R4, int R5, int R6)
        {
            int index = 2;
            ep.Workbook.Worksheets.Add($"{bayName}_{sheetName}");
            ExcelWorksheet st = ep.Workbook.Worksheets[$"{bayName}_{sheetName}"];
            st.Cells[1, 1].LoadFromText($"ID,TIMESTAMP,RG1,RP1,RG2,RP2,RG_B,RP_B");
            st.Column(2).Style.Numberformat.Format = @"yyyy/MM/dd HH:mm:ss.000";
            DateTime Temp = currentTime.AddHours(-32).AddMinutes(-1);
            for (int i = 0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint={R1} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0)) break;
                            st.Cells[index++, 1].LoadFromText($"{i + 1},{reader.GetDateTime(1).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")},{reader.GetDouble(0).ToString("C2")}");
                        }
                }
            }
            index = 2;
            for (int i = 0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value from analogevents where points_idPoint={R2} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0)) break;
                            st.Cells[index++, 4].LoadFromText($"{reader.GetDouble(0).ToString("C2")}");
                        }
                }
            }
            index = 2;
            for (int i = 0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value from analogevents where points_idPoint={R3} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0)) break;
                            st.Cells[index++, 5].LoadFromText($"{reader.GetDouble(0).ToString("C2")}");
                            //st.Cells[index++, 5].LoadFromText("NO DATA");
                        }
                }
            }
            index = 2;
            for (int i = 0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value from analogevents where points_idPoint={R4} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0)) break;
                            st.Cells[index++, 6].LoadFromText($"{reader.GetDouble(0).ToString("C2")}");
                            //st.Cells[index++, 6].LoadFromText("NO DATA");
                        }
                }
            }
            index = 2;
            for (int i = 0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value from analogevents where points_idPoint={R5} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0)) break;
                            //st.Cells[index++, 7].LoadFromText($"{reader.GetDouble(0)}");
                            st.Cells[index++, 7].LoadFromText("NO DATA");
                        }
                }
            }
            index = 2;
            for (int i = 0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value from analogevents where points_idPoint={R6} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0)) break;
                            //st.Cells[index++, 8].LoadFromText($"{reader.GetDouble(0)}");
                            st.Cells[index++, 8].LoadFromText("NO DATA");
                        }
                }
            }
            st.Cells.AutoFitColumns();
        }
        private void AddEnvironmentSheets(MySqlConnection connection, ExcelPackage ep, string bayName, string sheetName, DateTime currentTime, int WS, int WD, int Temp0, int RH)
        {
            int index = 2;
            ep.Workbook.Worksheets.Add($"{bayName}_{sheetName}");
            ExcelWorksheet st = ep.Workbook.Worksheets[$"{bayName}_{sheetName}"];
            st.Cells[1, 1].LoadFromText($"ID,TIMESTAMP,WS,WD,Temp,RH");
            st.Column(2).Style.Numberformat.Format = @"yyyy/MM/dd HH:mm:ss.000";
            DateTime Temp = currentTime.AddHours(-32).AddMinutes(-1);
            for (int i = 0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint={WS} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0)) break;
                            st.Cells[index++, 1].LoadFromText($"{i + 1},{reader.GetDateTime(1).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")},{reader.GetDouble(0).ToString("C2")}");
                        }
                }
            }
            index = 2;
            for (int i = 0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value from analogevents where points_idPoint={WD} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0)) break;
                            st.Cells[index++, 4].LoadFromText($"{reader.GetDouble(0).ToString("C2")}");
                        }
                }
            }
            index = 2;
            for (int i = 0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value from analogevents where points_idPoint={Temp0} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0)) break;
                            st.Cells[index++, 5].LoadFromText($"{reader.GetDouble(0).ToString("C2")}");
                        }
                }
            }
            index = 2;
            for (int i = 0; i < 1442; i++)
            {
                using (var command = new MySqlCommand($"select value from analogevents where points_idPoint={RH} and eventTime between '{Temp.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{Temp.AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                {
                    command.CommandTimeout = 6000;
                    using (var reader = command.ExecuteReader())
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0)) break;
                            st.Cells[index++, 6].LoadFromText($"{reader.GetDouble(0).ToString("C2")}");
                        }
                }
            }
            st.Cells.AutoFitColumns();
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

                oldState_TenDays = raiseFlag_TenDays;
                switch (currentTime.Day)
                {
                    case 1:
                        raiseState_TenDays = 0;
                        raiseFlag_TenDays = true;
                        break;
                    case 11:
                        raiseState_TenDays = 1;
                        raiseFlag_TenDays = true;
                        break;
                    case 21:
                        raiseState_TenDays = 2;
                        raiseFlag_TenDays = true;
                        break;
                    default:
                        raiseFlag_TenDays = false;
                        break;
                }
                
                oldState_Archive = raiseFlag_Archive;
                raiseFlag_Archive = currentTime.Hour == _options.Value.ArchiveTime ? true : false;

                DebugStr = DebugMode ? "_Debug" : "";

                //TenDays Report
                if ((oldState_TenDays == false && raiseFlag_TenDays == true && ArchiveOnly == false))
                {
                    DateTime temp;
                    temp = currentTime.AddDays(-10);
                    if (raiseState_TenDays == 0)
                    {
                        temp = new DateTime(temp.Year, temp.Month, 21, 0, 0, 0);
                    }
                    else
                    {
                        temp = new DateTime(temp.Year, temp.Month, temp.Day, 0, 0, 0);
                    }
                    if (!File.Exists($"{ _options.Value.ReportDirectory}\\Report\\Excel\\TenDays\\CHENHUA_TaipowerMonthlyReport_{temp.Month}月.xlsx"))
                    {
                        File.Copy(@"Template/CHENYA_TaipowerMonthlyTemplate.xlsx",
                            $"{ _options.Value.ReportDirectory}\\Report\\Excel\\TenDays\\CHENHUA_TaipowerMonthlyReport_{temp.Month}月.xlsx");
                    }
                    /*if (!File.Exists($"{ _options.Value.ReportDirectory}\\Report\\Excel\\TenDays\\HOLDGOOD_TaipowerMonthlyReport_{temp.Month}月.xlsx"))
                    {
                        File.Copy(@"Template/HOLDGOOD_TaipowerMonthlyTemplate.xlsx",
                            $"{ _options.Value.ReportDirectory}\\Report\\Excel\\TenDays\\HOLDGOOD_TaipowerMonthlyReport_{temp.Month}月.xlsx");
                    }*/
                    _logger.LogInformation($"Query DB to Excel File...");
                    try
                    {
                        using (ExcelPackage ep = new ExcelPackage(new FileInfo($"{ _options.Value.ReportDirectory}\\Report\\Excel\\TenDays\\CHENHUA_TaipowerMonthlyReport_{temp.Month}月.xlsx")))
                        {
                            using (var connection = new MySqlConnection($"Server={_options.Value.MySQL_IpAddress};User ID={_options.Value.MySQL_User};Password={_options.Value.MySQL_Password};Database={_options.Value.MySQL_DbTable}"))
                            {
                                connection.Open();
                                //Do something
                                string old_ws_name = "", ws_name = " ";
                                switch (raiseState_TenDays)
                                {
                                    case 0:
                                        old_ws_name = "X月下";
                                        ws_name = $"{temp.Month}月下";
                                        break;
                                    case 1:
                                        old_ws_name = "X月上";
                                        ws_name = $"{temp.Month}月上";
                                        break;
                                    case 2:
                                        old_ws_name = "X月中";
                                        ws_name = $"{temp.Month}月中";
                                        break;
                                }
                                var ws = ep.Workbook.Worksheets.SingleOrDefault(x => x.Name == old_ws_name);
                                // Init for WorkSheets
                                ws.Name = ws_name;
                                if (raiseState_TenDays == 0)
                                {
                                    for (int i = 1; i < (DateTime.DaysInMonth(temp.Year, temp.Month) - 20) * 2 + 1; i++)
                                    {
                                        ws.Cells[i + 6, 2].Value = temp.Year - 1911;
                                        ws.Cells[i + 6, 6].Value = temp.Month;
                                        ws.Cells[i + 6, 8].Value = 21 + ((i - 1) / 2);
                                    }
                                }
                                else
                                {
                                    for (int i = 1; i < 21; i++)
                                    {
                                        ws.Cells[i + 6, 2].Value = temp.Year - 1911;
                                        ws.Cells[i + 6, 6].Value = temp.Month;
                                    }
                                }
                                DateTime temp_UTC = temp.AddHours(-8);
                                int row_format = 0;
                                for (int DayOfMonth = 0; DayOfMonth < 11; DayOfMonth++)
                                {
                                    if ((raiseState_TenDays == 1 || raiseState_TenDays == 2) && DayOfMonth == 10)
                                    {
                                        break;
                                    }
                                    else if (DayOfMonth == (DateTime.DaysInMonth(temp.Year, temp.Month) - 20) && raiseState_TenDays == 0)
                                    {
                                        break;
                                    }
                                    double sum = 0;
                                    for (int HourOfDay = 0; HourOfDay < 24; HourOfDay++)
                                    {
                                        double Last_Value = 0, First_Value = 0;
                                        using (var command = new MySqlCommand($"select value from analogevents where points_idPoint=171 and eventTime " +
                                            $"between '{temp_UTC.AddDays(DayOfMonth).AddHours(HourOfDay).ToString("yyyy-MM-dd HH:00")}' and '{temp_UTC.AddDays(DayOfMonth).AddHours(HourOfDay + 1).ToString("yyyy-MM-dd HH:00")}' order by eventTime desc limit 1", connection))
                                        {
                                            command.CommandTimeout = 6000;
                                            using (var reader = command.ExecuteReader())
                                                while (reader.Read())
                                                {
                                                    Last_Value = reader.GetDouble(0);
                                                }
                                        }
                                        using (var command = new MySqlCommand($"select value from analogevents where points_idPoint=171 and eventTime " +
                                            $"between '{temp_UTC.AddDays(DayOfMonth).AddHours(HourOfDay).ToString("yyyy-MM-dd HH:00")}' and '{temp_UTC.AddDays(DayOfMonth).AddHours(HourOfDay + 1).ToString("yyyy-MM-dd HH:00")}' order by eventTime asc limit 1", connection))
                                        {
                                            command.CommandTimeout = 6000;
                                            using (var reader = command.ExecuteReader())
                                                while (reader.Read())
                                                {
                                                    First_Value = reader.GetDouble(0);
                                                }
                                        }
                                        if (HourOfDay < 5 || HourOfDay > 17)
                                        {
                                            if (HourOfDay == 23)
                                            {
                                                ws.Cells[7 + DayOfMonth + row_format + (HourOfDay / 12), 19].Value = sum;
                                                ws.Cells[7 + DayOfMonth + row_format + (HourOfDay / 12), 14 + (HourOfDay % 12)].Value = Last_Value - First_Value;
                                            }
                                            else
                                            {
                                                sum = sum + (Last_Value - First_Value);
                                                ws.Cells[7 + DayOfMonth + row_format + (HourOfDay / 12), 14 + (HourOfDay % 12)].Value = 0;
                                            }
                                        }
                                        else
                                        {
                                            if (HourOfDay == 17)
                                            {
                                                sum = 0;
                                            }
                                            if (HourOfDay == 5)
                                            {
                                                ws.Cells[7 + DayOfMonth + row_format + (HourOfDay / 12), 19].Value = sum;
                                            }
                                            else
                                            {
                                                ws.Cells[7 + DayOfMonth + row_format + (HourOfDay / 12), 14 + (HourOfDay % 12)].Value = Last_Value - First_Value;
                                            }
                                        }

                                        _logger.LogInformation($"{temp_UTC.AddDays(DayOfMonth).AddHours(HourOfDay).AddHours(8).ToString("yyyy-MM-dd HH:00")} - LV: {Last_Value}, FV: {First_Value}, Diff: {Last_Value - First_Value}");
                                    }
                                    row_format++;
                                }
                                ep.Save();
                            }
                        }
                        /*
                        using (ExcelPackage ep = new ExcelPackage(new FileInfo($"{ _options.Value.ReportDirectory}\\Report\\Excel\\TenDays\\HOLDGOOD_TaipowerMonthlyReport_{temp.Month}月.xlsx")))
                        {
                            using (var connection = new MySqlConnection($"Server={_options.Value.MySQL_IpAddress};User ID={_options.Value.MySQL_User};Password={_options.Value.MySQL_Password};Database={_options.Value.MySQL_DbTable}"))
                            {
                                connection.Open();
                                //Do something
                                string old_ws_name = "", ws_name = " ";
                                switch (raiseState_TenDays)
                                {
                                    case 0:
                                        old_ws_name = "X月下";
                                        ws_name = $"{temp.Month}月下";
                                        break;
                                    case 1:
                                        old_ws_name = "X月上";
                                        ws_name = $"{temp.Month}月上";
                                        break;
                                    case 2:
                                        old_ws_name = "X月中";
                                        ws_name = $"{temp.Month}月中";
                                        break;
                                }
                                var ws = ep.Workbook.Worksheets.SingleOrDefault(x => x.Name == old_ws_name);
                                // Init for WorkSheets
                                ws.Name = ws_name;
                                if (raiseState_TenDays == 0)
                                {
                                    for (int i = 1; i < (DateTime.DaysInMonth(temp.Year, temp.Month) - 20) * 2 + 1; i++)
                                    {
                                        ws.Cells[i + 6, 2].Value = temp.Year - 1911;
                                        ws.Cells[i + 6, 6].Value = temp.Month;
                                        ws.Cells[i + 6, 8].Value = 21 + ((i - 1) / 2);
                                    }
                                }
                                else
                                {
                                    for (int i = 1; i < 21; i++)
                                    {
                                        ws.Cells[i + 6, 2].Value = temp.Year - 1911;
                                        ws.Cells[i + 6, 6].Value = temp.Month;
                                    }
                                }
                                DateTime temp_UTC = temp.AddHours(-8);
                                int row_format = 0;
                                for (int DayOfMonth = 0; DayOfMonth < 11; DayOfMonth++)
                                {
                                    if ((raiseState_TenDays == 1 || raiseState_TenDays == 2) && DayOfMonth == 10)
                                    {
                                        break;
                                    }
                                    else if (DayOfMonth == (DateTime.DaysInMonth(temp.Year, temp.Month) - 20) && raiseState_TenDays == 0)
                                    {
                                        break;
                                    }
                                    double sum = 0;
                                    for (int HourOfDay = 0; HourOfDay < 24; HourOfDay++)
                                    {
                                        double Last_Value = 0, First_Value = 0;
                                        using (var command = new MySqlCommand($"select value from analogevents where points_idPoint=1068 and eventTime " +
                                            $"between '{temp_UTC.AddDays(DayOfMonth).AddHours(HourOfDay).ToString("yyyy-MM-dd HH:00")}' and '{temp_UTC.AddDays(DayOfMonth).AddHours(HourOfDay + 1).ToString("yyyy-MM-dd HH:00")}' order by eventTime desc limit 1", connection))
                                        {
                                            command.CommandTimeout = 6000;
                                            using (var reader = command.ExecuteReader())
                                                while (reader.Read())
                                                {
                                                    Last_Value = reader.GetDouble(0);
                                                }
                                        }
                                        using (var command = new MySqlCommand($"select value from analogevents where points_idPoint=1068 and eventTime " +
                                            $"between '{temp_UTC.AddDays(DayOfMonth).AddHours(HourOfDay).ToString("yyyy-MM-dd HH:00")}' and '{temp_UTC.AddDays(DayOfMonth).AddHours(HourOfDay + 1).ToString("yyyy-MM-dd HH:00")}' order by eventTime asc limit 1", connection))
                                        {
                                            command.CommandTimeout = 6000;
                                            using (var reader = command.ExecuteReader())
                                                while (reader.Read())
                                                {
                                                    First_Value = reader.GetDouble(0);
                                                }
                                        }

                                        if (HourOfDay < 5 || HourOfDay > 17)
                                        {
                                            if (HourOfDay == 23)
                                            {
                                                ws.Cells[7 + DayOfMonth + row_format + (HourOfDay / 12), 19].Value = sum;
                                                ws.Cells[7 + DayOfMonth + row_format + (HourOfDay / 12), 14 + (HourOfDay % 12)].Value = Last_Value - First_Value;
                                            }
                                            else
                                            {
                                                sum = sum + (Last_Value - First_Value);
                                                ws.Cells[7 + DayOfMonth + row_format + (HourOfDay / 12), 14 + (HourOfDay % 12)].Value = 0;
                                            }
                                        }
                                        else
                                        {
                                            if (HourOfDay == 17)
                                            {
                                                sum = 0;
                                            }
                                            if (HourOfDay == 5)
                                            {
                                                ws.Cells[7 + DayOfMonth + row_format + (HourOfDay / 12), 19].Value = sum;
                                            }
                                            else
                                            {
                                                ws.Cells[7 + DayOfMonth + row_format + (HourOfDay / 12), 14 + (HourOfDay % 12)].Value = Last_Value - First_Value;
                                            }
                                        }
                                    }
                                    row_format++;
                                }
                                ep.Save();
                            }
                        }*/
                    }
                    catch (Exception e)
                    {
                        _logger.LogError(e.Message);
                    }
                    
                }

                //Daily Report
                if ((oldState_Daily == false && raiseFlag_Daily == true && ArchiveOnly == false))
                {
                    _logger.LogInformation($"Query DB to Excel File...");
                    try
                    {
                        using (ExcelPackage ep = new ExcelPackage())
                        {
                            using (var connection = new MySqlConnection($"Server={_options.Value.MySQL_IpAddress};User ID={_options.Value.MySQL_User};Password={_options.Value.MySQL_Password};Database={_options.Value.MySQL_DbTable}"))
                            {
                                connection.Open();
                                /*
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
                                */
                                
                                AddDailySheets(connection, ep, "DTR1670", "REV", currentTime, 165);
                                AddDailySheets(connection, ep, "DTR1670", "FWD", currentTime, 167);
                                AddDailySheets(connection, ep, "MP5", "FWD", currentTime, 169);
                                AddDailySheets(connection, ep, "MP5", "REV", currentTime, 171);
                                AddDailySheets(connection, ep, "MP6", "FWD", currentTime, 173);
                                AddDailySheets(connection, ep, "MP6", "REV", currentTime, 175);
                                AddDailySheets(connection, ep, "FEEDER_31", "FWD", currentTime, 177);
                                AddDailySheets(connection, ep, "FEEDER_31", "REV", currentTime, 179);
                                AddDailySheets(connection, ep, "FEEDER_32", "FWD", currentTime, 181);
                                AddDailySheets(connection, ep, "FEEDER_32", "REV", currentTime, 183);
                                AddDailySheets(connection, ep, "FEEDER_33", "FWD", currentTime, 185);
                                AddDailySheets(connection, ep, "FEEDER_33", "REV", currentTime, 187);
                                AddDailySheets(connection, ep, "FEEDER_34", "FWD", currentTime, 189);
                                AddDailySheets(connection, ep, "FEEDER_34", "REV", currentTime, 191);
                                AddDailySheets(connection, ep, "FEEDER_35", "FWD", currentTime, 193);
                                AddDailySheets(connection, ep, "FEEDER_35", "REV", currentTime, 195);
                                AddDailySheets(connection, ep, "FEEDER_36", "FWD", currentTime, 197);
                                AddDailySheets(connection, ep, "FEEDER_36", "REV", currentTime, 199);
                                AddDeviceTempSheets(connection, ep, "海上01&02", "模組溫度", currentTime, 2092, 2100);
                                AddRadiationSheets(connection, ep, "海上01&02&棧橋", "水平垂直日照", currentTime, 2089, 2090, 2097, 2098, 2089, 2090);
                                AddEnvironmentSheets(connection, ep, "海上01", "環境數值", currentTime, 2094, 2095, 2091, 2096);
                                AddEnvironmentSheets(connection, ep, "海上02", "環境數值", currentTime, 2102, 2103, 2099, 2104);
                                //AddEnvironmentSheets(connection, ep, "棧橋", "環境數值", currentTime, 2094, 2095, 2091, 2096);

                                //EVENTLIST
                                int index = 2;
                                ep.Workbook.Worksheets.Add("EVENTLIST");
                                ExcelWorksheet EVENTLIST = ep.Workbook.Worksheets["EVENTLIST"];
                                EVENTLIST.Cells[1, 1].LoadFromText($"EVENT,STATE,TIMESTAMP");
                                EVENTLIST.Column(3).Style.Numberformat.Format = @"yyyy/MM/dd HH:mm:ss.000";
                                using (var command = new MySqlCommand($"select T3, state, eventTime from digitalevents inner join points on digitalevents.points_idPoint = points.idPoint where eventTime between '{currentTime.AddHours(-32).ToString("yyyy-MM-dd HH:mm:ss")}' and '{currentTime.AddHours(-8).ToString("yyyy-MM-dd HH:mm:ss")}'", connection))
                                using (var reader = command.ExecuteReader())
                                    while (reader.Read())
                                        if(reader.GetString(0) != "") EVENTLIST.Cells[index++, 1].LoadFromText($"{reader.GetString(0)},{reader.GetString(1)},{reader.GetDateTime(2).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")}");
                                EVENTLIST.Cells.AutoFitColumns();
                                /*
                                index = 2;
                                ep.Workbook.Worksheets.Add("SUNSHINEMETER_1");
                                ExcelWorksheet SUNSHINE_METER = ep.Workbook.Worksheets["SUNSHINEMETER_1"];
                                SUNSHINE_METER.Cells[1, 1].LoadFromText($"TIMESTAMP,SUNSHINE");
                                SUNSHINE_METER.Column(1).Style.Numberformat.Format = @"yyyy/MM/dd HH:mm:ss.000";

                                for (int i = 0; i < 1442; i++)
                                {
                                    using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=40 and eventTime between '{currentTime.AddHours(-32).AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{currentTime.AddHours(-32).AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                                    {
                                        command.CommandTimeout = 6000;
                                        using (var reader = command.ExecuteReader())
                                            while (reader.Read())
                                            {
                                                if (reader.IsDBNull(0)) break;
                                                SUNSHINE_METER.Cells[index++, 1].LoadFromText($"{reader.GetDateTime(1).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")},{reader.GetDouble(0)}");
                                            }
                                    }
                                }
                                SUNSHINE_METER.Cells.AutoFitColumns();

                                index = 2;
                                ep.Workbook.Worksheets.Add("SUNSHINEMETER_2");
                                ExcelWorksheet SUNSHINE_METER2 = ep.Workbook.Worksheets["SUNSHINEMETER_2"];
                                SUNSHINE_METER2.Cells[1, 1].LoadFromText($"TIMESTAMP,SUNSHINE");
                                SUNSHINE_METER2.Column(1).Style.Numberformat.Format = @"yyyy/MM/dd HH:mm:ss.000";

                                for (int i = 0; i < 1442; i++)
                                {
                                    using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=5767 and eventTime between '{currentTime.AddHours(-32).AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{currentTime.AddHours(-32).AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                                    {
                                        command.CommandTimeout = 6000;
                                        using (var reader = command.ExecuteReader())
                                            while (reader.Read())
                                            {
                                                if (reader.IsDBNull(0)) break;
                                                SUNSHINE_METER2.Cells[index++, 1].LoadFromText($"{reader.GetDateTime(1).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss.fff")},{reader.GetDouble(0)}");
                                            }
                                    }
                                }
                                SUNSHINE_METER2.Cells.AutoFitColumns();
                                */
                                //Save ExcelFile
                                FileInfo fi = new FileInfo($"{_options.Value.ReportDirectory}\\Report\\Excel\\Daily\\CHENHUA-{currentTime.AddDays(-1).ToString("yyyyMMdd")}_DailyReport{DebugStr}.xlsx");
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
                if ((oldState_Monthly == false && raiseFlag_Monthly == true && ArchiveOnly == false) || DebugMode)
                {
                    using (ExcelPackage ep = new ExcelPackage())
                    {
                        using (var connection = new MySqlConnection($"Server={_options.Value.MySQL_IpAddress};User ID={_options.Value.MySQL_User};Password={_options.Value.MySQL_Password};Database={_options.Value.MySQL_DbTable}"))
                        {
                            connection.Open();
                            /*
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
                            */

                            AddMonthlySheets(connection, ep, "DTR1670", "REV", currentTime, 165, true);
                            AddMonthlySheets(connection, ep, "DTR1670", "FWD", currentTime, 167, true);
                            AddMonthlySheets(connection, ep, "MP5", "FWD", currentTime, 169);
                            AddMonthlySheets(connection, ep, "MP5", "REV", currentTime, 171);
                            AddMonthlySheets(connection, ep, "MP6", "FWD", currentTime, 173);
                            AddMonthlySheets(connection, ep, "MP6", "REV", currentTime, 175);
                            AddMonthlySheets(connection, ep, "FEEDER_31", "FWD", currentTime, 177);
                            AddMonthlySheets(connection, ep, "FEEDER_31", "REV", currentTime, 179);
                            AddMonthlySheets(connection, ep, "FEEDER_32", "FWD", currentTime, 181);
                            AddMonthlySheets(connection, ep, "FEEDER_32", "REV", currentTime, 183);
                            AddMonthlySheets(connection, ep, "FEEDER_33", "FWD", currentTime, 185);
                            AddMonthlySheets(connection, ep, "FEEDER_33", "REV", currentTime, 187);
                            AddMonthlySheets(connection, ep, "FEEDER_34", "FWD", currentTime, 189);
                            AddMonthlySheets(connection, ep, "FEEDER_34", "REV", currentTime, 191);
                            AddMonthlySheets(connection, ep, "FEEDER_35", "FWD", currentTime, 193);
                            AddMonthlySheets(connection, ep, "FEEDER_35", "REV", currentTime, 195);
                            AddMonthlySheets(connection, ep, "FEEDER_36", "FWD", currentTime, 197);
                            AddMonthlySheets(connection, ep, "FEEDER_36", "REV", currentTime, 199);


                            FileInfo fi = new FileInfo($"{_options.Value.ReportDirectory}\\Report\\Excel\\Monthly\\CHENHUA-{currentTime.AddMonths(-1).ToString("yyyyMM")}_MonthlyReport{DebugStr}.xlsx");
                            ep.SaveAs(fi);
                        }
                    }

                    //Simple MonthlyReport
                    DateTime temp = currentTime.AddMonths(-1);
                    temp = new DateTime(temp.Year, temp.Month, 1);
                    if (!File.Exists($"{ _options.Value.ReportDirectory}\\Report\\Excel\\Monthly\\{temp.Month}月簡明月報.xlsx"))
                    {
                        File.Copy(@"Template/Simple_MonthlyReportTemplate.xlsx",
                            $"{ _options.Value.ReportDirectory}\\Report\\Excel\\Monthly\\{temp.Month}月簡明月報.xlsx");
                    }
                    using (ExcelPackage ep = new ExcelPackage(new FileInfo($"{ _options.Value.ReportDirectory}\\Report\\Excel\\Monthly\\{temp.Month}月簡明月報.xlsx")))
                    {
                        using (var connection = new MySqlConnection($"Server={_options.Value.MySQL_IpAddress};User ID={_options.Value.MySQL_User};Password={_options.Value.MySQL_Password};Database={_options.Value.MySQL_DbTable}"))
                        {
                            connection.Open();
                            var ws = ep.Workbook.Worksheets.SingleOrDefault(x => x.Name == "辰華電力");
                            // Init for WorkSheets
                            DateTime temp_UTC = temp.AddHours(-8);
                            double Last_Value = 0, First_Value = 0;
                            for (int DayOfMonth = 0; DayOfMonth < DateTime.DaysInMonth(temp.Year, temp.Month); DayOfMonth++)
                            {
                                ws.Cells[2, 3 + DayOfMonth].Value = temp.AddDays(DayOfMonth).ToString(@"MM/dd");

                                //單日發電量
                                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=171 and eventTime " +
                                                $"between '{temp_UTC.AddDays(DayOfMonth).ToString("yyyy-MM-dd 16:00:00")}' and '{temp_UTC.AddDays(DayOfMonth + 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime desc limit 1", connection))
                                {
                                    command.CommandTimeout = 6000;
                                    using (var reader = command.ExecuteReader())
                                        while (reader.Read())
                                        {
                                            Last_Value = reader.GetDouble(0);
                                        }
                                }
                                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=171 and eventTime " +
                                    $"between '{temp_UTC.AddDays(DayOfMonth).ToString("yyyy-MM-dd 16:00:00")}' and '{temp_UTC.AddDays(DayOfMonth + 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime asc limit 1", connection))
                                {
                                    command.CommandTimeout = 6000;
                                    using (var reader = command.ExecuteReader())
                                        while (reader.Read())
                                        {
                                            First_Value = reader.GetDouble(0);
                                        }
                                }

                                ws.Cells[3, 3 + DayOfMonth].Value = Last_Value - First_Value;
                                Last_Value = 0; First_Value = 0;
                                //單日耗電量
                                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=169 and eventTime " +
                                                $"between '{temp_UTC.AddDays(DayOfMonth).ToString("yyyy-MM-dd 16:00:00")}' and '{temp_UTC.AddDays(DayOfMonth + 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime desc limit 1", connection))
                                {
                                    command.CommandTimeout = 6000;
                                    using (var reader = command.ExecuteReader())
                                        while (reader.Read())
                                        {
                                            Last_Value = reader.GetDouble(0);
                                        }
                                }
                                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=169 and eventTime " +
                                    $"between '{temp_UTC.AddDays(DayOfMonth).ToString("yyyy-MM-dd 16:00:00")}' and '{temp_UTC.AddDays(DayOfMonth + 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime asc limit 1", connection))
                                {
                                    command.CommandTimeout = 6000;
                                    using (var reader = command.ExecuteReader())
                                        while (reader.Read())
                                        {
                                            First_Value = reader.GetDouble(0);
                                        }
                                }

                                ws.Cells[4, 3 + DayOfMonth].Value = Last_Value - First_Value;
                                Last_Value = 0; First_Value = 0;
                                /*
                                //所內耗電量
                                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=1022 and eventTime " +
                                                $"between '{temp_UTC.AddDays(DayOfMonth).ToString("yyyy-MM-dd 16:00:00")}' and '{temp_UTC.AddDays(DayOfMonth + 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime desc limit 1", connection))
                                {
                                    command.CommandTimeout = 6000;
                                    using (var reader = command.ExecuteReader())
                                        while (reader.Read())
                                        {
                                            Last_Value = reader.GetDouble(0);
                                        }
                                }
                                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=1022 and eventTime " +
                                    $"between '{temp_UTC.AddDays(DayOfMonth).ToString("yyyy-MM-dd 16:00:00")}' and '{temp_UTC.AddDays(DayOfMonth + 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime asc limit 1", connection))
                                {
                                    command.CommandTimeout = 6000;
                                    using (var reader = command.ExecuteReader())
                                        while (reader.Read())
                                        {
                                            First_Value = reader.GetDouble(0);
                                        }
                                }

                                ws.Cells[5, 3 + DayOfMonth].Value = Last_Value - First_Value;
                                */
                                bool startFlag = false, endFlag = false;
                                double last_value = 0;
                                DateTime startTime = new DateTime(2000, 1, 1, 0, 0, 0), endTime = new DateTime(2000, 1, 1, 0, 0, 0);
                                for (int i = 0; i < 1442; i++)
                                {
                                    using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=171 and eventTime between '{temp_UTC.AddDays(DayOfMonth).AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{temp_UTC.AddDays(DayOfMonth).AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                                    {
                                        command.CommandTimeout = 6000;
                                        using (var reader = command.ExecuteReader())
                                            while (reader.Read())
                                            {
                                                if (endFlag) break;
                                                if (i == 0)
                                                {
                                                    last_value = reader.GetDouble(0);
                                                    break;
                                                }
                                                if (last_value == reader.GetDouble(0) && !startFlag)
                                                {
                                                    break;
                                                }
                                                if (last_value == reader.GetDouble(0) && startFlag)
                                                {
                                                    endTime = reader.GetDateTime(1);
                                                    endFlag = true;
                                                    break;
                                                }
                                                if (!startFlag)
                                                {
                                                    startTime = reader.GetDateTime(1);
                                                    startFlag = true;
                                                }
                                                last_value = reader.GetDouble(0);
                                            }
                                    }
                                }
                                ws.Cells[6, 3 + DayOfMonth].Value = startTime.AddHours(8).ToString("HH:mm");
                                ws.Cells[7, 3 + DayOfMonth].Value = endTime.AddHours(8).ToString("HH:mm");
                                ws.Cells[8, 3 + DayOfMonth].Value = $"{endTime.Subtract(startTime).Hours}:{endTime.Subtract(startTime).Minutes}";
                            }
                            ep.Save();
                        }
                    }
                    /*
                    using (ExcelPackage ep = new ExcelPackage(new FileInfo($"{ _options.Value.ReportDirectory}\\Report\\Excel\\Monthly\\{temp.Month}月簡明月報.xlsx")))
                    {
                        using (var connection = new MySqlConnection($"Server={_options.Value.MySQL_IpAddress};User ID={_options.Value.MySQL_User};Password={_options.Value.MySQL_Password};Database={_options.Value.MySQL_DbTable}"))
                        {
                            connection.Open();
                            var ws = ep.Workbook.Worksheets.SingleOrDefault(x => x.Name == "厚固光電");
                            // Init for WorkSheets
                            DateTime temp_UTC = temp.AddHours(-8);
                            double Last_Value = 0, First_Value = 0;
                            for (int DayOfMonth = 0; DayOfMonth < DateTime.DaysInMonth(temp.Year, temp.Month); DayOfMonth++)
                            {
                                ws.Cells[2, 3 + DayOfMonth].Value = temp.AddDays(DayOfMonth).ToString(@"MM/dd");

                                //單日發電量
                                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=1068 and eventTime " +
                                                $"between '{temp_UTC.AddDays(DayOfMonth).ToString("yyyy-MM-dd 16:00:00")}' and '{temp_UTC.AddDays(DayOfMonth + 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime desc limit 1", connection))
                                {
                                    command.CommandTimeout = 6000;
                                    using (var reader = command.ExecuteReader())
                                        while (reader.Read())
                                        {
                                            Last_Value = reader.GetDouble(0);
                                        }
                                }
                                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=1068 and eventTime " +
                                    $"between '{temp_UTC.AddDays(DayOfMonth).ToString("yyyy-MM-dd 16:00:00")}' and '{temp_UTC.AddDays(DayOfMonth + 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime asc limit 1", connection))
                                {
                                    command.CommandTimeout = 6000;
                                    using (var reader = command.ExecuteReader())
                                        while (reader.Read())
                                        {
                                            First_Value = reader.GetDouble(0);
                                        }
                                }

                                ws.Cells[3, 3 + DayOfMonth].Value = Last_Value - First_Value;
                                Last_Value = 0; First_Value = 0;

                                //單日耗電量
                                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=1070 and eventTime " +
                                                $"between '{temp_UTC.AddDays(DayOfMonth).ToString("yyyy-MM-dd 16:00:00")}' and '{temp_UTC.AddDays(DayOfMonth + 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime desc limit 1", connection))
                                {
                                    command.CommandTimeout = 6000;
                                    using (var reader = command.ExecuteReader())
                                        while (reader.Read())
                                        {
                                            Last_Value = reader.GetDouble(0);
                                        }
                                }
                                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=1070 and eventTime " +
                                    $"between '{temp_UTC.AddDays(DayOfMonth).ToString("yyyy-MM-dd 16:00:00")}' and '{temp_UTC.AddDays(DayOfMonth + 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime asc limit 1", connection))
                                {
                                    command.CommandTimeout = 6000;
                                    using (var reader = command.ExecuteReader())
                                        while (reader.Read())
                                        {
                                            First_Value = reader.GetDouble(0);
                                        }
                                }

                                ws.Cells[4, 3 + DayOfMonth].Value = Last_Value - First_Value;
                                Last_Value = 0; First_Value = 0;

                                //所內耗電量
                                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=1022 and eventTime " +
                                                $"between '{temp_UTC.AddDays(DayOfMonth).ToString("yyyy-MM-dd 16:00:00")}' and '{temp_UTC.AddDays(DayOfMonth + 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime desc limit 1", connection))
                                {
                                    command.CommandTimeout = 6000;
                                    using (var reader = command.ExecuteReader())
                                        while (reader.Read())
                                        {
                                            Last_Value = reader.GetDouble(0);
                                        }
                                }
                                using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=1022 and eventTime " +
                                    $"between '{temp_UTC.AddDays(DayOfMonth).ToString("yyyy-MM-dd 16:00:00")}' and '{temp_UTC.AddDays(DayOfMonth + 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime asc limit 1", connection))
                                {
                                    command.CommandTimeout = 6000;
                                    using (var reader = command.ExecuteReader())
                                        while (reader.Read())
                                        {
                                            First_Value = reader.GetDouble(0);
                                        }
                                }

                                ws.Cells[5, 3 + DayOfMonth].Value = Last_Value - First_Value;

                                bool startFlag = false, endFlag = false;
                                double last_value = 0;
                                DateTime startTime = new DateTime(2000, 1, 1, 0, 0, 0), endTime = new DateTime(2000, 1, 1, 0, 0, 0);
                                for (int i = 0; i < 1442; i++)
                                {
                                    using (var command = new MySqlCommand($"select value, eventTime from analogevents where points_idPoint=1068 and eventTime between '{temp_UTC.AddDays(DayOfMonth).AddMinutes(i).ToString("yyyy-MM-dd HH:mm:00")}' and '{temp_UTC.AddDays(DayOfMonth).AddMinutes(i + 1).ToString("yyyy-MM-dd HH:mm:00")}' order by eventTime desc limit 1", connection))
                                    {
                                        command.CommandTimeout = 6000;
                                        using (var reader = command.ExecuteReader())
                                            while (reader.Read())
                                            {
                                                if (endFlag) break;
                                                if (i == 0)
                                                {
                                                    last_value = reader.GetDouble(0);
                                                    break;
                                                }
                                                if (last_value == reader.GetDouble(0) && !startFlag)
                                                {
                                                    break;
                                                }
                                                if (last_value == reader.GetDouble(0) && startFlag)
                                                {
                                                    endTime = reader.GetDateTime(1);
                                                    endFlag = true;
                                                    break;
                                                }
                                                if (!startFlag)
                                                {
                                                    startTime = reader.GetDateTime(1);
                                                    startFlag = true;
                                                }
                                                last_value = reader.GetDouble(0);
                                            }
                                    }
                                }
                                ws.Cells[6, 3 + DayOfMonth].Value = startTime.AddHours(8).ToString("HH:mm");
                                ws.Cells[7, 3 + DayOfMonth].Value = endTime.AddHours(8).ToString("HH:mm");
                                ws.Cells[8, 3 + DayOfMonth].Value = $"{endTime.Subtract(startTime).Hours}:{endTime.Subtract(startTime).Minutes}";
                            }
                            ep.Save();
                        }
                    }*/
                }
                //Archive Service
                if(oldState_Archive == false && raiseFlag_Archive == true && Archive_Is_Finished == true && ReportOnly == false)
                {
                    Archive_Is_Finished = false;
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
                            using (var scan_end = new MySqlCommand($"select idEvent from analogevents where eventTime < '{QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd 16:00:00")}' and eventTime > '{QueryDate.AddDays(AddDay - 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime desc limit 1", connection))
                            {
                                scan_end.CommandTimeout = 6000;
                                using (var result_end = scan_end.ExecuteReader())
                                    while (result_end.Read())
                                    {
                                        TodayHaveRawdata = true;
                                        ScanIndex_End = result_end.GetInt32(0);
                                    }
                            }
                            _logger.LogInformation($"DATE {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} scan completed.");
                            //insert log to DB
                            InsertMsgToDbTable("DATABASE SCAN END", $"DATE: {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} scan completed.");
                            if (TodayHaveRawdata)
                            {
                                using (var scan_start = new MySqlCommand($"select idEvent from analogevents where eventTime < '{QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd 16:00:00")}' and eventTime > '{QueryDate.AddDays(AddDay - 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime asc limit 1", connection))
                                {
                                    scan_start.CommandTimeout = 6000;
                                    using (var result_start = scan_start.ExecuteReader())
                                        while (result_start.Read())
                                        {
                                            ScanIndex_Start = result_start.GetInt32(0);
                                        }
                                }
                            }
                            else
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
                                string ori_path = $"{_options.Value.BackupDirectory}\\{QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")}_Rawdata";
                                string path = ori_path;
                                int file_count = 0;
                                while (File.Exists($"{path}.zip"))
                                {
                                    path = $"{ori_path}_{file_count}";
                                    file_count++;
                                }
                                //Compress and save SQL file to *.zip
                                using (FileStream zipToOpen = new FileStream($"{path}.zip", FileMode.Create))
                                {
                                    using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                                    {
                                        ZipArchiveEntry readmeEntry = archive.CreateEntry($"{QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")}_Rawdata.sql");
                                        using (StreamWriter writer = new StreamWriter(readmeEntry.Open()))
                                        {
                                            ProcessStartInfo proc = new ProcessStartInfo();
                                            string cmd = $"--skip-tz-utc --no-create-info --host={_options.Value.MySQL_IpAddress} --user={_options.Value.MySQL_User} --password={_options.Value.MySQL_Password} {_options.Value.MySQL_DbTable} analogevents --where=\"eventTime < '{QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd 16:00:00")}' and eventTime > '{QueryDate.AddDays(AddDay - 1).ToString("yyyy-MM-dd 15:59:59")}'\"";
                                            // Configure path for mysqldump.exe
                                            proc.FileName = _options.Value.EXEPATH;
                                            proc.RedirectStandardInput = false;
                                            proc.RedirectStandardOutput = true;
                                            proc.UseShellExecute = false;
                                            proc.WindowStyle = ProcessWindowStyle.Minimized;
                                            proc.Arguments = cmd;
                                            proc.CreateNoWindow = true;
                                            Process p = Process.Start(proc);
                                            writer.Write(p.StandardOutput.ReadToEnd());
                                            p.WaitForExit();
                                            p.Close();
                                        }
                                    }
                                }

                                TodayHaveRawdata = false;
                                _logger.LogInformation($"SaveFile: \"{path}.zip\"");
                                _logger.LogInformation($"DATE {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} archive completed.");
                                InsertMsgToDbTable("EXPORT DATA END", $"DATE: {path}.zip export completed and saving file.");

                                FileInfo fileinfo = new FileInfo($"{path}.zip");
                                if (fileinfo.Length < 10000)
                                {
                                    break;
                                }

                                //Remove rawdata from database
                                _logger.LogInformation($"Remove \"{QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")}\" rawdata from database.");
                                InsertMsgToDbTable("REMOVE DATA START", $"DATE: {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} removing old data.");
                                using (var command = new MySqlCommand($"delete from analogevents where eventTime < '{QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd 16:00:00")}' and eventTime > '{QueryDate.AddDays(AddDay - 1).ToString("yyyy-MM-dd 15:59:59")}' order by eventTime", connection))
                                {
                                    command.CommandTimeout = 604800;
                                    command.ExecuteNonQuery();
                                }
                                _logger.LogInformation($"Remove completed.");
                                InsertMsgToDbTable("REMOVE DATA END", $"DATE: {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} remove completed.");

                            }
                            catch (Exception e)
                            {
                                _logger.LogError($"Error occured when exporting data: {e.Message}");
                                InsertMsgToDbTable("ERROR OCCURED WHEN EXPORTING", $"DATE: {QueryDate.AddDays(AddDay).ToString("yyyy-MM-dd")} export failed.");
                                continue;
                            }
                            AddDay--;
                        }
                    }
                    Archive_Is_Finished = true;
                }
                await Task.Delay(1000, stoppingToken);
            }
        }
    }
}
