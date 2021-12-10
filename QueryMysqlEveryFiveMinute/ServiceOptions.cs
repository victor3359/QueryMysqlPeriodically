using System;
using System.Collections.Generic;
using System.Text;

namespace QueryMysqlEveryFiveMinute
{
    public class ServiceOptions
    {
        public string ReportDirectory { get; set; }
        public string MySQL_IpAddress { get; set; }
        public string MySQL_User { get; set; }
        public string MySQL_Password { get; set; }
        public string MySQL_DbTable { get; set; }
        public string BackupDirectory { get; set; }
        public string EXEPATH { get; set; }
        public int ArchiveTime { get; set; }
        public int DailyReportTime { get; set; }
        public int KeepDataAliveDays { get; set; }
    }
}
